VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{9B8DFCFF-6726-11D6-99D4-000244193CE4}#1.0#0"; "PROJECT1.OCX"
Begin VB.Form FRM_ENDOSCOPY 
   Caption         =   "Endoscopy"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
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
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      ForeColor       =   &H8000000D&
      Height          =   3255
      Left            =   0
      TabIndex        =   52
      Top             =   3480
      Width           =   11775
      Begin VB.OptionButton opt_eddiagnostic 
         Caption         =   "Diagnostic"
         Height          =   255
         Left            =   6120
         TabIndex        =   72
         Top             =   1440
         Width           =   1335
      End
      Begin VB.OptionButton opt_edtherapetic 
         Caption         =   "Therapeutic"
         Height          =   195
         Left            =   6120
         TabIndex        =   71
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6360
         TabIndex        =   68
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6600
         TabIndex        =   66
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         TabIndex        =   64
         Top             =   2640
         Width           =   2895
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         TabIndex        =   62
         Top             =   2160
         Width           =   2895
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         TabIndex        =   60
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         TabIndex        =   58
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         TabIndex        =   56
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox text1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   54
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label8 
         Caption         =   "On research project (if any)"
         Height          =   495
         Left            =   4800
         TabIndex        =   67
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Additional Remarks"
         Height          =   375
         Left            =   5160
         TabIndex        =   65
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Immed. Complications"
         Height          =   375
         Left            =   120
         TabIndex        =   63
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Biopsy/Cytology"
         Height          =   375
         Left            =   120
         TabIndex        =   61
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Modifiers"
         Height          =   375
         Left            =   120
         TabIndex        =   59
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Primary/Secondary Diagnosis"
         Height          =   375
         Left            =   120
         TabIndex        =   57
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Site Seen  && Normal"
         Height          =   375
         Left            =   120
         TabIndex        =   55
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Site && Finding"
         Height          =   375
         Left            =   240
         TabIndex        =   53
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame COMMAND 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      TabIndex        =   43
      Top             =   7800
      Width           =   11775
      Begin VB.CommandButton CMD_NEW 
         Caption         =   "&New"
         Height          =   375
         Left            =   360
         MouseIcon       =   "frm_endoscopy.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   51
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_view 
         Caption         =   "&View"
         Height          =   375
         Left            =   1800
         MouseIcon       =   "frm_endoscopy.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   50
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_modify 
         Caption         =   "&Modify"
         Height          =   375
         Left            =   3120
         MouseIcon       =   "frm_endoscopy.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   49
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_save 
         Caption         =   "&Save"
         Height          =   375
         Left            =   4440
         MouseIcon       =   "frm_endoscopy.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   48
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   5880
         MouseIcon       =   "frm_endoscopy.frx":0C28
         MousePointer    =   99  'Custom
         TabIndex        =   47
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "&Print"
         Height          =   375
         Left            =   8760
         MouseIcon       =   "frm_endoscopy.frx":0F32
         MousePointer    =   99  'Custom
         TabIndex        =   46
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_back 
         Caption         =   "&Back"
         Height          =   375
         Left            =   10320
         MouseIcon       =   "frm_endoscopy.frx":123C
         MousePointer    =   99  'Custom
         TabIndex        =   45
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_commit 
         Caption         =   "&Commit"
         Height          =   375
         Left            =   7200
         MouseIcon       =   "frm_endoscopy.frx":1546
         MousePointer    =   99  'Custom
         TabIndex        =   44
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      TabIndex        =   38
      Top             =   6720
      Width           =   11775
      Begin VB.TextBox TXT_COMMENTS 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   69
         Top             =   240
         Width           =   10215
      End
      Begin VB.ComboBox cmb_enterby 
         Height          =   315
         Left            =   1080
         TabIndex        =   40
         Top             =   600
         Width           =   1935
      End
      Begin VB.ComboBox cmb_finalby 
         Height          =   315
         Left            =   9480
         TabIndex        =   39
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label LBL_COMM 
         AutoSize        =   -1  'True
         Caption         =   "Remarks"
         Height          =   195
         Left            =   240
         TabIndex        =   70
         Top             =   240
         Width           =   750
      End
      Begin VB.Label lbl_enter 
         AutoSize        =   -1  'True
         Caption         =   "Enter by"
         Height          =   195
         Left            =   240
         TabIndex        =   42
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lbl_finby 
         AutoSize        =   -1  'True
         Caption         =   "To be Finalized by"
         Height          =   195
         Left            =   7680
         TabIndex        =   41
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Frame common 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   11775
      Begin VB.ComboBox cmb_modality 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6240
         TabIndex        =   18
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txt_refdoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   17
         Top             =   2160
         Width           =   3855
      End
      Begin VB.TextBox txt_modcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10440
         TabIndex        =   16
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txt_bedno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7920
         TabIndex        =   15
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txt_ward 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4920
         TabIndex        =   14
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txt_pridoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   13
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txt_status 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10440
         TabIndex        =   12
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txt_category 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7920
         TabIndex        =   11
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txt_age 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5520
         TabIndex        =   10
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txt_sex 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   9
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txt_name 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   8
         Top             =   1080
         Width           =   9135
      End
      Begin VB.TextBox txt_pathno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10200
         TabIndex        =   7
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox cmb_reqno 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6240
         TabIndex        =   6
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txt_deptcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10200
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmb_dept 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2400
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   0
         TabIndex        =   1
         Top             =   720
         Width           =   1335
         Begin VB.OptionButton opt_staff 
            Caption         =   "Staff"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton Opt_patient 
            Caption         =   "Patient"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   1095
         End
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   2400
         TabIndex        =   19
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   0
         PromptInclude   =   0   'False
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "??>/#####"
         PromptChar      =   "_"
      End
      Begin VB.Label lbl_refdoc 
         AutoSize        =   -1  'True
         Caption         =   "Referring Doctor/service"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   2160
         Width           =   2130
      End
      Begin VB.Label lbl_Modalitycode 
         AutoSize        =   -1  'True
         Caption         =   "Modality Code"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9120
         TabIndex        =   34
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lbl_modality 
         AutoSize        =   -1  'True
         Caption         =   "Modality"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5400
         TabIndex        =   33
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl_bed 
         AutoSize        =   -1  'True
         Caption         =   "Bed No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6960
         TabIndex        =   32
         Top             =   1800
         Width           =   705
      End
      Begin VB.Label lbl_ward 
         Caption         =   "Ward"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4320
         TabIndex        =   31
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lbl_primarydoc 
         AutoSize        =   -1  'True
         Caption         =   "Primary Doctor/Service"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   1800
         Width           =   1995
      End
      Begin VB.Label lbl_status 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9720
         TabIndex        =   29
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label lbl_category 
         AutoSize        =   -1  'True
         Caption         =   "Category"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6960
         TabIndex        =   28
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label lbl_age 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5040
         TabIndex        =   27
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label lbl_sex 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1920
         TabIndex        =   26
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label lbl_name 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1680
         TabIndex        =   25
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lbl_pathno 
         AutoSize        =   -1  'True
         Caption         =   "Lab No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9360
         TabIndex        =   24
         Top             =   600
         Width           =   690
      End
      Begin VB.Label lbl_reqno 
         AutoSize        =   -1  'True
         Caption         =   "Requisition No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4920
         TabIndex        =   23
         Top             =   720
         Width           =   1320
      End
      Begin VB.Label lbl_caseno 
         AutoSize        =   -1  'True
         Caption         =   "Case No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1440
         TabIndex        =   22
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lbl_deptcode 
         AutoSize        =   -1  'True
         Caption         =   "Dept. Code"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9120
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbl_dept 
         AutoSize        =   -1  'True
         Caption         =   "Requisition Placed to"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   20
         Top             =   360
         Width           =   1830
      End
   End
   Begin Project1.UserControl1 UserControl11 
      Height          =   615
      Left            =   4200
      TabIndex        =   36
      Top             =   0
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1085
   End
   Begin VB.Label lbl_biochemlab 
      AutoSize        =   -1  'True
      Caption         =   "ENDOSCOPY DATA COLLECTION"
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
      Left            =   5280
      TabIndex        =   37
      Top             =   600
      Width           =   4695
   End
End
Attribute VB_Name = "FRM_ENDOSCOPY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_back_Click()
'GO BACK
DIS_MAIN.show
Me.Hide

End Sub

Private Sub cmd_cancel_Click()
'CLEAR ALL FIELD
End Sub

Private Sub cmd_commit_Click()
'COMMIT PRESENT RECORD
End Sub

Private Sub cmd_modify_Click()
'ALLOW TO MODIFY THE SELECTED RECORD
End Sub

Private Sub CMD_NEW_Click()
'ALLOW TO ENTER NEW RECORD
End Sub

Private Sub cmd_print_Click()
'GET THE PRINT OUT FOR THIS RECORD
End Sub

Private Sub cmd_save_Click()
'SAVE PRESENT RECORD AND CHANGE THE BUTON STATE
End Sub

Private Sub cmd_view_Click()
'GET THE RECORD FOR VIEW ONLY
End Sub

Private Sub lbl_site3_Click()
End Sub

Private Sub tab_endo_DblClick()

End Sub
