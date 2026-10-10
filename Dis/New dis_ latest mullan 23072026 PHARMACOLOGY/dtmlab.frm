VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form dtmlab 
   Caption         =   "DTM lab Requisition"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12900
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
   ScaleHeight     =   5625
   ScaleWidth      =   12900
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_name 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5400
      TabIndex        =   13
      Top             =   1680
      Width           =   5655
   End
   Begin VB.CommandButton CMD_BACK 
      Caption         =   "<< &Back"
      Height          =   495
      Left            =   9600
      TabIndex        =   11
      Top             =   7680
      Width           =   1335
   End
   Begin VB.TextBox txt_others 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   10
      Top             =   4770
      Width           =   5055
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1935
      Left            =   840
      TabIndex        =   0
      Top             =   2280
      Width           =   10095
      Begin VB.CheckBox chk_HCV 
         Caption         =   "Anti - HCV(ELISA)"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   1440
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CheckBox chk_h015 
         Caption         =   "HIV Western Blot"
         Height          =   255
         Left            =   3840
         TabIndex        =   17
         Top             =   960
         Width           =   2295
      End
      Begin VB.CheckBox CHK_AS 
         Caption         =   "Antibody Screen"
         Height          =   255
         Left            =   7200
         TabIndex        =   12
         Top             =   480
         Width           =   2415
      End
      Begin VB.CheckBox CHK_DAT 
         Caption         =   "Direct Antiglobulin Test"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3840
         TabIndex        =   8
         Top             =   480
         Width           =   2415
      End
      Begin VB.CheckBox CHK_IAT 
         Caption         =   "Indirect Antiglobulin Test"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Top             =   960
         Width           =   2535
      End
      Begin VB.CheckBox chk_BG 
         Caption         =   "Blood Group"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   720
         TabIndex        =   5
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox chk_HBSAG 
         Caption         =   "HBsAg(ELISA)"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4080
         TabIndex        =   4
         Top             =   1440
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox chk_HIV 
         Caption         =   "Anti - HIV( ELISA)"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5880
         TabIndex        =   2
         Top             =   1440
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CheckBox chk_cardenz 
         Caption         =   "Investigation "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   2055
      End
   End
   Begin MSMask.MaskEdBox msk_caseno 
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      ForeColor       =   16711680
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
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
      Left            =   720
      TabIndex        =   18
      Top             =   0
      Width           =   10335
   End
   Begin VB.Label lbl_caseno 
      Caption         =   "Case No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1320
      TabIndex        =   16
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lbl_name 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4680
      TabIndex        =   15
      Top             =   1680
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Others"
      Height          =   195
      Left            =   1320
      TabIndex        =   9
      Top             =   4800
      Width           =   570
   End
   Begin VB.Label Label2 
      Caption         =   "DTM LAB - REQUISITION"
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
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   600
      Width           =   3615
   End
End
Attribute VB_Name = "dtmlab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub chk_cardenz_Click()
If chk_cardenz.Value = Checked Then
    chk_BG.Value = Checked
    chk_HBSAG.Value = Checked
    chk_HCV.Value = Checked
    chk_HIV.Value = Checked
Else
    chk_BG.Value = Unchecked
    chk_HBSAG.Value = Unchecked
    chk_HCV.Value = Unchecked
    chk_HIV.Value = Unchecked
End If
End Sub

Private Sub cmd_back_Click()
BG1 = ""
ccount = 0
pathology.CMB_TEXT.clear
If DTMFLAG = False Then
    If chk_BG.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        pathology.CMB_TEXT.AddItem "H001"
        BG1 = "Y"
    End If
    'sundeep  by dr.Narayan
   ' If chk_HBSAG.Value = 1 Then
   '     If ccount >= 11 Then
   '         MsgBox " You Can Select Max. of 11 Procedures only."
   '         Unload Me
   '         pathology.show
   '         Exit Sub
   '     End If
    '    ccount = ccount + 1
   '     pathology.CMB_TEXT.AddItem "H012"
   ' End If
   'sundeep  by dr.Narayan
   
   'sundeep  by dr.Narayan
    'If chk_HCV.Value = 1 Then
    '    If ccount >= 11 Then
    '        MsgBox " You Can Select Max. of 11 Procedures only."
    '        Unload Me
    '        pathology.show
    '        Exit Sub
    '    End If
    '    ccount = ccount + 1
    '    pathology.CMB_TEXT.AddItem "H014"
    'End If
    'sundeep  by dr.Narayan
    
    'sundeep  by dr.Narayan
    'If chk_HIV.Value = 1 Then
    '    If ccount >= 11 Then
    '        MsgBox " You Can Select Max. of 11 Procedures only."
    '        Unload Me
    '        pathology.show
     '       Exit Sub
     '   End If
     '   ccount = ccount + 1
     '   pathology.CMB_TEXT.AddItem "H013"
   ' End If
   'sundeep  by dr.Narayan
    If CHK_DAT.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        pathology.CMB_TEXT.AddItem "H006"
    End If
    If CHK_IAT.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        pathology.CMB_TEXT.AddItem "H007"
    End If
    If CHK_AS.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        pathology.CMB_TEXT.AddItem "?"
    End If
    If chk_h015.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        pathology.CMB_TEXT.AddItem "H015"
    End If
    HTOther = ""
    If Trim(TXT_OTHERS.Text) <> "" Then
        HTOther = Trim(TXT_OTHERS.Text)
    End If
    
End If
DTMFLAG = False
Unload Me
pathology.show
End Sub


