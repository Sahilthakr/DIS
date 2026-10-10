VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form hlatyping 
   Caption         =   "HLA Typing"
   ClientHeight    =   9705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11745
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
   ScaleHeight     =   9705
   ScaleWidth      =   11745
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chk_abdr 
      Caption         =   "HLA AB-DR (Molecular Typing) "
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5880
      TabIndex        =   24
      Top             =   5880
      Width           =   3135
   End
   Begin VB.TextBox txt_age 
      Height          =   285
      Left            =   8520
      TabIndex        =   11
      Top             =   3600
      Width           =   855
   End
   Begin VB.ComboBox cmb_sex 
      Height          =   315
      Left            =   4440
      TabIndex        =   10
      Top             =   3600
      Width           =   1335
   End
   Begin VB.ComboBox cmb_relation 
      Height          =   315
      Left            =   8520
      TabIndex        =   9
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox txt_donorname 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      TabIndex        =   8
      Top             =   3120
      Width           =   2535
   End
   Begin VB.CheckBox chk_donor 
      Caption         =   "Donor"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox txt_name 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4440
      TabIndex        =   2
      Top             =   2520
      Width           =   5655
   End
   Begin VB.CommandButton cmd_back 
      Caption         =   "<< Back"
      Height          =   615
      Left            =   9960
      TabIndex        =   18
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Investigations required"
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
      Height          =   3135
      Left            =   600
      TabIndex        =   1
      Top             =   4440
      Width           =   10575
      Begin VB.CheckBox chk_h227 
         Caption         =   "HLA -C && DQ (Molecular Typing) "
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5280
         TabIndex        =   25
         Top             =   1920
         Width           =   3135
      End
      Begin VB.CheckBox CHK_201 
         Caption         =   "HLA - DNA Typing DNA - DRB1 Low Resolution "
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   1560
         Width           =   5175
      End
      Begin VB.CheckBox CHK_101 
         Caption         =   "HLA Typing - ABC And DR "
         Height          =   375
         Left            =   600
         TabIndex        =   13
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox txt_others 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Top             =   2520
         Width           =   3135
      End
      Begin VB.CheckBox chk_102 
         Caption         =   "HLA  Typing - Only ABC"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   600
         TabIndex        =   12
         Top             =   480
         Width           =   3975
      End
      Begin VB.CheckBox chk_103 
         Caption         =   "HLA - Typing - Only DR"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5280
         TabIndex        =   15
         Top             =   480
         Width           =   2655
      End
      Begin VB.CheckBox chk_104 
         Caption         =   "HLA Typing - B-27"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5280
         TabIndex        =   16
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Others"
         Height          =   195
         Left            =   600
         TabIndex        =   6
         Top             =   2520
         Width           =   570
      End
   End
   Begin MSMask.MaskEdBox msk_caseno 
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   2520
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
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   23
      Top             =   480
      Width           =   10815
   End
   Begin VB.Label Label6 
      Caption         =   "Age"
      Height          =   255
      Left            =   7800
      TabIndex        =   22
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Sex"
      Height          =   375
      Left            =   3840
      TabIndex        =   21
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Relation"
      Height          =   255
      Left            =   7440
      TabIndex        =   20
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Name"
      Height          =   255
      Left            =   3600
      TabIndex        =   19
      Top             =   3120
      Width           =   855
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
      Left            =   600
      TabIndex        =   5
      Top             =   2520
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
      Left            =   3720
      TabIndex        =   4
      Top             =   2520
      Width           =   630
   End
   Begin VB.Label Label2 
      Caption         =   "HLA TYPING - REQUISITION"
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
      Left            =   3840
      TabIndex        =   0
      Top             =   1200
      Width           =   4215
   End
End
Attribute VB_Name = "hlatyping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub chk_donor_Click()
If chk_donor.Value Then
    cmb_relation.AddItem "Mother"
    cmb_relation.AddItem "Father"
    cmb_relation.AddItem "Brother"
    cmb_relation.AddItem "Sister"
    cmb_relation.AddItem "Daughter"
    cmb_relation.AddItem "Son"
    cmb_relation.AddItem "Other"
    txt_donorname.SetFocus
    Exit Sub
End If
End Sub

Private Sub cmb_relation_LostFocus()
If chk_donor.Value And Trim(cmb_relation.Text) = "" Then
    MsgBox "Pls Enter Relation !!", vbInformation, "HLA Typing"
    'cmb_relation.SetFocus
    Exit Sub
End If
End Sub

Private Sub cmd_back_Click()
ccount = 0
pathology.CMB_TEXT.clear
'If HLAFLAG = False Then
    If CHK_101.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        pathology.CMB_TEXT.AddItem "H101"
    End If
    If chk_102.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        pathology.CMB_TEXT.AddItem "H102"
    End If
    If chk_103.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        pathology.CMB_TEXT.AddItem "H103"
    End If
    If chk_104.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        pathology.CMB_TEXT.AddItem "H104"
    End If
    If CHK_201.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        pathology.CMB_TEXT.AddItem "H201"
    End If
    If chk_abdr.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        pathology.CMB_TEXT.AddItem "H226"
    End If
    
    ''''new test is added by shweta 29-sept-2012
    
        If chk_h227.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        pathology.CMB_TEXT.AddItem "H227"
    End If
    
    
'End If
    HTOther = ""
    DonorRelation = ""
    meterial = ""
    STRDonor = ""
    HLA_SEX = ""
    HLA_AGE = ""
    
If Trim(txt_donorname.Text) <> "" And chk_donor.Value Then
    meterial = Trim(txt_donorname.Text)
    DonorRelation = Trim(cmb_relation.Text)
    STRDonor = "yes"
    HLA_AGE = Trim(txt_age.Text)
    HLA_SEX = Trim(cmb_sex.Text)
End If
If Trim(txt_others.Text) <> "" Then
    HTOther = Trim(txt_others.Text)
End If
'If Trim(txt_age.Text) <> "" Then
'    HLA_AGE = Trim(txt_age.Text)
'End If
'If Trim(cmb_sex.Text) <> "" Then
'    HLA_SEX = Trim(cmb_sex.Text)
'End If
If chk_donor.Value And Trim(txt_donorname.Text) = "" And Trim(cmb_relation.Text) = "" And Trim(cmb_sex.Text) = "" And Trim(txt_age.Text) = "" Then
    MsgBox "Pls Enter Donor Details !!", vbInformation, "HLA TYPING"
    Exit Sub
End If
Unload Me
pathology.show
End Sub

Private Sub Form_Load()
cmb_sex.AddItem "M"
cmb_sex.AddItem "F"
End Sub

Private Sub txt_donorname_LostFocus()
If chk_donor.Value And Trim(txt_donorname.Text) = "" Then
    MsgBox "Pls Enter Donor Name !!", vbInformation, "HLA Typing"
    'txt_donorname.SetFocus
    Exit Sub
End If
End Sub
