VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frm_SubPath 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Sub Form of Pathology"
   ClientHeight    =   11100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16290
   LinkTopic       =   "Form2"
   ScaleHeight     =   11100
   ScaleWidth      =   16290
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMD_SAVE1 
      BackColor       =   &H00808080&
      Caption         =   "&Save"
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   10560
      Width           =   2175
   End
   Begin VB.CommandButton cmd_PRINT1 
      BackColor       =   &H00808080&
      Caption         =   "&Print"
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
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   10560
      Width           =   2175
   End
   Begin RichTextLib.RichTextBox txt_Impres 
      Height          =   1095
      Left            =   2640
      TabIndex        =   11
      Top             =   9360
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   1931
      _Version        =   393217
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frm_SubPath.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txt_microscopic 
      Height          =   2655
      Left            =   2640
      TabIndex        =   10
      Top             =   6600
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   4683
      _Version        =   393217
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frm_SubPath.frx":0085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txt_Frozen 
      Height          =   1695
      Left            =   2640
      TabIndex        =   9
      Top             =   1680
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   2990
      _Version        =   393217
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frm_SubPath.frx":010A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txt_gross 
      Height          =   3015
      Left            =   2640
      TabIndex        =   8
      Top             =   3480
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   5318
      _Version        =   393217
      ScrollBars      =   3
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frm_SubPath.frx":018F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmd_back 
      BackColor       =   &H00808080&
      Caption         =   "&Commit"
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
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   10560
      Width           =   2175
   End
   Begin VB.TextBox txt_microscopic1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      MaxLength       =   7000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "frm_SubPath.frx":0214
      Top             =   6720
      Width           =   1815
   End
   Begin VB.TextBox txt_gross1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3480
      MaxLength       =   8000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "frm_SubPath.frx":021A
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox txt_Frozen1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      MaxLength       =   2000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "frm_SubPath.frx":0220
      Top             =   1680
      Width           =   2175
   End
   Begin RichTextLib.RichTextBox txt_nmat 
      Height          =   735
      Left            =   2640
      TabIndex        =   16
      Top             =   840
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   1296
      _Version        =   393217
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frm_SubPath.frx":0226
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Nature of Material"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   15
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   615
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   14775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Impression"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   9360
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Microscopic Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   6600
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Gross Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Frozen Section Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   2295
   End
End
Attribute VB_Name = "frm_SubPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_back_Click()

SPPATHREPORT.FLDENABL1 ("MODIFY")

SPPATHREPORT.txt_grossdesc.Text = Trim(txt_gross.Text)
SPPATHREPORT.txt_microdesc.Text = Trim(txt_microscopic.Text)
SPPATHREPORT.txt_frosecrep.Text = Trim(txt_Frozen.Text)
SPPATHREPORT.txt_natmeterial.Text = Trim(txt_nmat.Text)

'SPPATHREPORT.cmd_save_Click
SPPATHREPORT.cmd_commit_Click
'SPPATHREPORT.cmd_print_Click
Me.Hide
SPPATHREPORT.show
End Sub

Private Sub CMD_PRINT1_Click()
SPPATHREPORT.FLDENABL1 ("MODIFY")
SPPATHREPORT.txt_grossdesc.Text = Trim(txt_gross.Text)
SPPATHREPORT.txt_microdesc.Text = Trim(txt_microscopic.Text)
SPPATHREPORT.txt_frosecrep.Text = Trim(txt_Frozen.Text)
SPPATHREPORT.txt_grossdesc.Text = Trim(txt_gross.Text)
SPPATHREPORT.txt_microdesc.Text = Trim(txt_microscopic.Text)
SPPATHREPORT.txt_natmeterial.Text = Trim(txt_nmat.Text)
SPPATHREPORT.cmd_print_Click
Me.Hide
SPPATHREPORT.show

End Sub

Private Sub CMD_SAVE1_Click()
SPPATHREPORT.FLDENABL1 ("MODIFY")

SPPATHREPORT.txt_grossdesc.Text = Trim(txt_gross.Text)
SPPATHREPORT.txt_microdesc.Text = Trim(txt_microscopic.Text)
SPPATHREPORT.txt_frosecrep.Text = Trim(txt_Frozen.Text)
SPPATHREPORT.txt_natmeterial.Text = Trim(txt_nmat.Text)
SPPATHREPORT.txt_impfindiag.Text = Trim(txt_Impres.Text)

SPPATHREPORT.cmd_save_Click
'Me.Hide
'SPPATHREPORT.show

End Sub

Private Sub Form_Load()

'Read already entered data in text fields.
txt_gross.Text = Trim(SPPATHREPORT.txt_grossdesc.Text)
txt_microscopic.Text = SPPATHREPORT.txt_microdesc.Text
txt_Frozen.Text = SPPATHREPORT.txt_frosecrep.Text
txt_Impres.Text = SPPATHREPORT.txt_impfindiag.Text
Label5.Caption = "Case No : " & Trim(SPPATHREPORT.MaskEdBox1.FormattedText) & " Req No. : " & Trim(SPPATHREPORT.cmb_reqno.Text) & vbCrLf & " Patient Name : " & Trim(SPPATHREPORT.txt_name.Text)
SPPATHREPORT.FLDENABL1 ("VIEW")


End Sub

