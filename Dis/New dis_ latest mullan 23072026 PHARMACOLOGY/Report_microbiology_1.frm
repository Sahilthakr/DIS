VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form Report_microbiology1 
   Caption         =   "Microbiology"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   570
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
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame frm_bacteria 
      Appearance      =   0  'Flat
      Caption         =   "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   -240
      TabIndex        =   158
      Top             =   3120
      Width           =   12120
      Begin VB.ComboBox txt_comment 
         Height          =   315
         Left            =   2160
         TabIndex        =   225
         Top             =   4440
         Width           =   6495
      End
      Begin VB.ComboBox cmb_org3 
         Height          =   315
         Left            =   3240
         TabIndex        =   165
         Top             =   2400
         Width           =   7095
      End
      Begin VB.ComboBox cmb_org2 
         Height          =   315
         Left            =   3240
         TabIndex        =   164
         Top             =   2040
         Width           =   7095
      End
      Begin VB.ComboBox cmb_org1 
         Height          =   315
         Left            =   3240
         TabIndex        =   163
         Top             =   1680
         Width           =   7095
      End
      Begin VB.TextBox txt_report 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   1680
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   212
         Top             =   3720
         Visible         =   0   'False
         Width           =   7575
      End
      Begin VB.TextBox txt_epith 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   171
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txt_pus 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3960
         MaxLength       =   30
         TabIndex        =   172
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txt_other 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7320
         MaxLength       =   250
         TabIndex        =   173
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txt_ziehl 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10080
         MaxLength       =   15
         TabIndex        =   174
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txt_koh 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   176
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox cmb_gpc 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "Report_microbiology_1.frx":0000
         Left            =   3960
         List            =   "Report_microbiology_1.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   178
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox cmb_gnc 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "Report_microbiology_1.frx":002F
         Left            =   7320
         List            =   "Report_microbiology_1.frx":003C
         Style           =   2  'Dropdown List
         TabIndex        =   180
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox cmb_gpb 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "Report_microbiology_1.frx":005E
         Left            =   10080
         List            =   "Report_microbiology_1.frx":006B
         Style           =   2  'Dropdown List
         TabIndex        =   182
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox cmb_gnb 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "Report_microbiology_1.frx":008D
         Left            =   1560
         List            =   "Report_microbiology_1.frx":009A
         Style           =   2  'Dropdown List
         TabIndex        =   184
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox cmb_patho 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "Report_microbiology_1.frx":00BC
         Left            =   1560
         List            =   "Report_microbiology_1.frx":00C6
         Style           =   2  'Dropdown List
         TabIndex        =   170
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txt_organ1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3480
         TabIndex        =   169
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox TXT_ORGAN2 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3480
         TabIndex        =   168
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox TXT_ORGAN3 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3480
         TabIndex        =   166
         Top             =   2400
         Width           =   2295
      End
      Begin VB.ComboBox cmb_india 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "Report_microbiology_1.frx":00D3
         Left            =   7320
         List            =   "Report_microbiology_1.frx":00D5
         Style           =   2  'Dropdown List
         TabIndex        =   186
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox cmb_antibiotic 
         Height          =   315
         Left            =   1800
         TabIndex        =   167
         Top             =   3120
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox txt_3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7680
         MaxLength       =   1
         TabIndex        =   162
         Top             =   3120
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txt_2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6360
         MaxLength       =   1
         TabIndex        =   161
         Top             =   3120
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txt_commentold 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1560
         MaxLength       =   160
         MultiLine       =   -1  'True
         TabIndex        =   160
         Top             =   4440
         Width           =   7575
      End
      Begin VB.TextBox txt_1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4920
         MaxLength       =   1
         TabIndex        =   159
         Top             =   3120
         Visible         =   0   'False
         Width           =   1140
      End
      Begin MSFlexGridLib.MSFlexGrid grid_anti 
         Height          =   375
         Left            =   4680
         TabIndex        =   175
         Top             =   1320
         Visible         =   0   'False
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   661
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         FixedCols       =   0
         ForeColor       =   0
         ForeColorFixed  =   128
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         Appearance      =   0
         FormatString    =   "   ANTIBIOTIC                               |^         1           |^          2         |^              3    "
      End
      Begin MSFlexGridLib.MSFlexGrid grid_dmbanti 
         Height          =   1455
         Left            =   1560
         TabIndex        =   177
         Top             =   2880
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   2566
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         ForeColor       =   0
         ForeColorFixed  =   128
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         Appearance      =   0
         FormatString    =   "   ANTIBIOTIC                               |^         1           |^          2         |^              3    "
      End
      Begin VB.Label lbl_3 
         AutoSize        =   -1  'True
         Caption         =   "3."
         DragMode        =   1  'Automatic
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3000
         TabIndex        =   199
         Top             =   2400
         Width           =   180
      End
      Begin VB.Label lbl_2 
         AutoSize        =   -1  'True
         Caption         =   "2."
         DragMode        =   1  'Automatic
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3000
         TabIndex        =   198
         Top             =   2040
         Width           =   180
      End
      Begin VB.Label lbl_1 
         AutoSize        =   -1  'True
         Caption         =   "1."
         DragMode        =   1  'Automatic
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3000
         TabIndex        =   197
         Top             =   1680
         Width           =   180
      End
      Begin VB.Label lbl_organisms 
         AutoSize        =   -1  'True
         Caption         =   "Organisms Isolated"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   196
         Top             =   1800
         Width           =   1620
      End
      Begin VB.Label lbl_Epith 
         AutoSize        =   -1  'True
         Caption         =   "Epithelial Cells"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   195
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label lbl_pus 
         AutoSize        =   -1  'True
         Caption         =   "Pus Cells"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3120
         TabIndex        =   194
         Top             =   240
         Width           =   795
      End
      Begin VB.Label lbl_other 
         AutoSize        =   -1  'True
         Caption         =   "Other Findings"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6000
         TabIndex        =   193
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label lbl_ziehl 
         AutoSize        =   -1  'True
         Caption         =   "Z. N. Stain"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9120
         TabIndex        =   192
         Top             =   240
         Width           =   945
      End
      Begin VB.Label lbl_koh 
         AutoSize        =   -1  'True
         Caption         =   "K O H"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   840
         TabIndex        =   191
         Top             =   600
         Width           =   525
      End
      Begin VB.Label lbl_gpc 
         AutoSize        =   -1  'True
         Caption         =   "GPC"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3480
         TabIndex        =   190
         Top             =   600
         Width           =   390
      End
      Begin VB.Label lbl_gnc 
         AutoSize        =   -1  'True
         Caption         =   "GNC"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6840
         TabIndex        =   189
         Top             =   600
         Width           =   405
      End
      Begin VB.Label lbl_gpb 
         AutoSize        =   -1  'True
         Caption         =   "GPB"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9600
         TabIndex        =   188
         Top             =   600
         Width           =   390
      End
      Begin VB.Label lbl_GNB 
         AutoSize        =   -1  'True
         Caption         =   "GNB"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   960
         TabIndex        =   187
         Top             =   960
         Width           =   405
      End
      Begin VB.Label lbl_india 
         Caption         =   "India Ink Preparation for Cryptococcus"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3960
         TabIndex        =   185
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Pathogen Grown"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   183
         Top             =   1320
         Width           =   1425
      End
      Begin VB.Label lbl_report 
         AutoSize        =   -1  'True
         Caption         =   "Report"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   960
         TabIndex        =   181
         Top             =   3960
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lbl_bcomments 
         Caption         =   "Remarks"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   179
         Top             =   4440
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4695
      Left            =   0
      TabIndex        =   200
      Top             =   3120
      Width           =   11895
      Begin VB.TextBox txt_bresult 
         Appearance      =   0  'Flat
         Height          =   855
         Left            =   3960
         MaxLength       =   3000
         MultiLine       =   -1  'True
         TabIndex        =   202
         Top             =   1920
         Width           =   5655
      End
      Begin VB.TextBox txt_btest 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3960
         TabIndex        =   201
         Top             =   960
         Width           =   5655
      End
      Begin VB.Label Label12 
         Caption         =   "Result"
         Height          =   255
         Left            =   1920
         TabIndex        =   204
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Test Performed"
         Height          =   255
         Left            =   2040
         TabIndex        =   203
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmd_decomit 
      BackColor       =   &H008080FF&
      Caption         =   "Decommit"
      Height          =   495
      Left            =   6840
      TabIndex        =   251
      Top             =   10200
      Width           =   1215
   End
   Begin VB.CommandButton cmd_fullreport 
      BackColor       =   &H000000FF&
      Caption         =   "EXIT"
      Height          =   495
      Left            =   5520
      TabIndex        =   249
      Top             =   10200
      Width           =   1215
   End
   Begin VB.CommandButton cmd_save 
      Caption         =   "&Save"
      Height          =   495
      Left            =   2760
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   9600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame COMMAND 
      Height          =   1455
      Left            =   -840
      TabIndex        =   59
      Top             =   9360
      Width           =   12615
      Begin VB.CommandButton CMD_DECOMMIT 
         BackColor       =   &H000000FF&
         Caption         =   "<< BACK"
         Height          =   495
         Left            =   5280
         MaskColor       =   &H000000FF&
         TabIndex        =   60
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmd_cancel 
         BackColor       =   &H000000FF&
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   6840
         MousePointer    =   99  'Custom
         TabIndex        =   211
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "&Print"
         Height          =   495
         Left            =   5760
         MousePointer    =   99  'Custom
         TabIndex        =   210
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmd_previous 
         Caption         =   "&Pre_req"
         Height          =   495
         Left            =   9120
         TabIndex        =   209
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmd_commit 
         Caption         =   "&Commit"
         Height          =   495
         Left            =   4680
         MousePointer    =   99  'Custom
         TabIndex        =   65
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmd_back 
         BackColor       =   &H000000FF&
         Caption         =   "& NEXT_req"
         Height          =   495
         Left            =   7920
         MaskColor       =   &H00E0E0E0&
         MousePointer    =   99  'Custom
         TabIndex        =   64
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_modify 
         Caption         =   "&Modify"
         Height          =   495
         Left            =   2400
         MousePointer    =   99  'Custom
         TabIndex        =   63
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmd_view 
         Caption         =   "&View"
         Height          =   495
         Left            =   1320
         MousePointer    =   99  'Custom
         TabIndex        =   62
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CMD_NEW 
         Caption         =   "&New"
         Height          =   495
         Left            =   120
         MousePointer    =   99  'Custom
         TabIndex        =   61
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10080
      TabIndex        =   248
      Top             =   10680
      Width           =   2415
   End
   Begin VB.TextBox txt_date 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9600
      TabIndex        =   67
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txt_time 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9600
      TabIndex        =   66
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame frm_enter 
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   0
      TabIndex        =   50
      Top             =   8160
      Width           =   12255
      Begin MSMask.MaskEdBox msk_coll_tm 
         Height          =   255
         Left            =   10680
         TabIndex        =   218
         Top             =   480
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_coll_dt 
         Height          =   255
         Left            =   10680
         TabIndex        =   217
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cmb_enteryby 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1320
         Sorted          =   -1  'True
         TabIndex        =   54
         Top             =   120
         Width           =   3135
      End
      Begin VB.ComboBox cmb_final 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5880
         Sorted          =   -1  'True
         TabIndex        =   53
         Top             =   120
         Width           =   3855
      End
      Begin VB.TextBox txt_addendum 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   52
         Top             =   480
         Width           =   3135
      End
      Begin VB.ComboBox cmb_addendumby 
         Height          =   315
         Left            =   5880
         TabIndex        =   51
         Top             =   480
         Width           =   3855
      End
      Begin MSMask.MaskEdBox msk_fromdt 
         Height          =   255
         Left            =   14520
         TabIndex        =   216
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   0
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label15 
         Caption         =   "24 Hrs"
         Height          =   255
         Left            =   11520
         TabIndex        =   221
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lbl_colltm 
         Caption         =   "Coll Tm :"
         Height          =   255
         Left            =   9840
         TabIndex        =   215
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lbl_colldt 
         Caption         =   "Coll Dt : "
         Height          =   255
         Left            =   9840
         TabIndex        =   214
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lbl_enterby 
         AutoSize        =   -1  'True
         Caption         =   "Entered By"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   58
         Top             =   120
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Finalised By"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4800
         TabIndex        =   57
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label lbl_addendum 
         Caption         =   "ADDENDUM"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Addendum By"
         Height          =   255
         Left            =   4680
         TabIndex        =   55
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame common 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   1.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      TabIndex        =   12
      Top             =   720
      Width           =   11775
      Begin VB.TextBox txt_microno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10320
         TabIndex        =   30
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox cmb_dept 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2400
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   120
         Width           =   2295
      End
      Begin VB.TextBox txt_deptcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   9720
         TabIndex        =   29
         Top             =   120
         Width           =   1335
      End
      Begin VB.ComboBox cmb_reqno 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5160
         TabIndex        =   34
         Text            =   "cmb_reqno"
         Top             =   480
         Width           =   4335
      End
      Begin VB.TextBox txt_name 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         TabIndex        =   28
         Top             =   960
         Width           =   6255
      End
      Begin VB.TextBox txt_age 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         TabIndex        =   27
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txt_category 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4440
         TabIndex        =   26
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txt_status 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6240
         TabIndex        =   25
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txt_pridoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         TabIndex        =   24
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txt_ward 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   8040
         TabIndex        =   23
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txt_bedno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   9720
         TabIndex        =   22
         Top             =   1320
         Width           =   975
      End
      Begin VB.ComboBox cmb_modality 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6360
         Sorted          =   -1  'True
         TabIndex        =   32
         Top             =   120
         Width           =   2295
      End
      Begin VB.TextBox txt_sex 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   9720
         TabIndex        =   21
         Top             =   840
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1335
         Begin VB.OptionButton opt_staff 
            Caption         =   "Staff"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton Opt_patient 
            Caption         =   "Patient"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox txt_modcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10800
         TabIndex        =   17
         Top             =   960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox CMB_PROC 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8280
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txt_refdoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6840
         TabIndex        =   15
         Top             =   1680
         Width           =   3855
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF80&
         Caption         =   "&EMR"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txt_spm 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         TabIndex        =   13
         Top             =   2040
         Width           =   2055
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   2400
         TabIndex        =   33
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   0
         PromptInclude   =   0   'False
         MaxLength       =   8
         Mask            =   "??>/#####"
         PromptChar      =   "_"
      End
      Begin VB.Label lbl_category 
         AutoSize        =   -1  'True
         Caption         =   "Category                    Status"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3600
         TabIndex        =   41
         Top             =   1320
         Width           =   2505
      End
      Begin VB.Label lbl_LABID 
         AutoSize        =   -1  'True
         Caption         =   "Micro No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9480
         TabIndex        =   49
         Top             =   480
         Width           =   840
      End
      Begin VB.Label lbl_dept 
         AutoSize        =   -1  'True
         Caption         =   "Requisition Placed to"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   480
         TabIndex        =   48
         Top             =   240
         Width           =   1830
      End
      Begin VB.Label lbl_deptcode 
         AutoSize        =   -1  'True
         Caption         =   "Dept. Code"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8760
         TabIndex        =   47
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lbl_caseno 
         AutoSize        =   -1  'True
         Caption         =   "Case No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1440
         TabIndex        =   46
         Top             =   600
         Width           =   795
      End
      Begin VB.Label lbl_reqno 
         AutoSize        =   -1  'True
         Caption         =   "Requisition No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3840
         TabIndex        =   45
         Top             =   600
         Width           =   1320
      End
      Begin VB.Label lbl_name 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1800
         TabIndex        =   44
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lbl_sex 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9240
         TabIndex        =   43
         Top             =   840
         Width           =   330
      End
      Begin VB.Label lbl_age 
         AutoSize        =   -1  'True
         Caption         =   "Age "
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1920
         TabIndex        =   42
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label lbl_primarydoc 
         AutoSize        =   -1  'True
         Caption         =   "Primary Doctor/Service"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   40
         Top             =   1680
         Width           =   1995
      End
      Begin VB.Label lbl_ward 
         AutoSize        =   -1  'True
         Caption         =   "Ward               Bed No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7560
         TabIndex        =   39
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label lbl_modality 
         AutoSize        =   -1  'True
         Caption         =   "Modality"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5520
         TabIndex        =   38
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lbl_proc 
         AutoSize        =   -1  'True
         Caption         =   "Procedure"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7200
         TabIndex        =   37
         Top             =   2040
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lbl_refdoc 
         AutoSize        =   -1  'True
         Caption         =   "Referring Doctor/service"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4680
         TabIndex        =   36
         Top             =   1680
         Width           =   2130
      End
      Begin VB.Label Label13 
         Caption         =   "Specimen"
         Height          =   255
         Left            =   3240
         TabIndex        =   35
         Top             =   2040
         Width           =   1695
      End
   End
   Begin Crystal.CrystalReport CRYSTALREPORT1 
      Left            =   2040
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   1440
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame frm_others 
      ForeColor       =   &H00C00000&
      Height          =   4335
      Left            =   0
      TabIndex        =   75
      Top             =   3120
      Width           =   11775
      Begin VB.ComboBox cmb_motemplet 
         Height          =   315
         Left            =   1920
         TabIndex        =   227
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txt_result 
         Appearance      =   0  'Flat
         Height          =   1005
         Left            =   1920
         MultiLine       =   -1  'True
         TabIndex        =   77
         Top             =   1920
         Width           =   9255
      End
      Begin VB.TextBox txt_test_perform 
         Appearance      =   0  'Flat
         Height          =   435
         Left            =   1920
         MultiLine       =   -1  'True
         TabIndex        =   76
         Top             =   720
         Width           =   9255
      End
      Begin VB.Label Label17 
         Caption         =   "TEMPLET"
         Height          =   255
         Left            =   480
         TabIndex        =   226
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lbl_result 
         AutoSize        =   -1  'True
         Caption         =   "Result"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1200
         TabIndex        =   79
         Top             =   2280
         Width           =   555
      End
      Begin VB.Label lbl_test 
         AutoSize        =   -1  'True
         Caption         =   "Test Performed"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   480
         TabIndex        =   78
         Top             =   840
         Width           =   1305
      End
   End
   Begin VB.Frame frm_myco 
      Caption         =   "MYCOBACTERIOLOGY"
      Height          =   4740
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   11535
      Begin VB.ComboBox cmb_antiboitic_mb 
         Height          =   315
         Left            =   1200
         TabIndex        =   222
         Top             =   2160
         Width           =   3735
      End
      Begin VB.TextBox txt_c_comment 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1200
         MaxLength       =   160
         TabIndex        =   5
         Top             =   3720
         Width           =   8655
      End
      Begin VB.ComboBox cmb_stain 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Text            =   "cmb_stain"
         Top             =   480
         Width           =   2655
      End
      Begin VB.ComboBox cmb_organisms 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4560
         TabIndex        =   3
         Top             =   465
         Width           =   6855
      End
      Begin VB.TextBox txt_c_report 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1200
         Width           =   8655
      End
      Begin VB.TextBox txt_c_1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4800
         TabIndex        =   1
         Top             =   2160
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid grid_c_anti 
         Height          =   1575
         Left            =   1200
         TabIndex        =   6
         Top             =   1920
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   2778
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         ForeColor       =   0
         ForeColorFixed  =   128
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         Appearance      =   0
         FormatString    =   "   ANTIMICROBIAL                               |^         1           "
      End
      Begin VB.Label lbl_c_comments 
         Caption         =   "Remarks"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label lbl_stain 
         AutoSize        =   -1  'True
         Caption         =   "Z. N. Stain"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Organism Isolated"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5520
         TabIndex        =   8
         Top             =   240
         Width           =   1530
      End
      Begin VB.Label lbl_c_report 
         Caption         =   " Provisional            Culture            Report"
         ForeColor       =   &H00000000&
         Height          =   675
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1275
      End
   End
   Begin VB.Frame frm_clinic 
      Height          =   4575
      Left            =   0
      TabIndex        =   80
      Top             =   3240
      Width           =   11655
      Begin VB.Frame frm_urineexam 
         Height          =   2235
         Left            =   0
         TabIndex        =   100
         Top             =   0
         Width           =   11535
         Begin VB.Frame frm_urine_micro 
            Caption         =   "Urine Microscopy Examination"
            ForeColor       =   &H000000C0&
            Height          =   2070
            Left            =   5760
            TabIndex        =   150
            Top             =   120
            Width           =   5655
            Begin VB.TextBox txt_urinecomment 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   960
               MaxLength       =   252
               TabIndex        =   124
               Top             =   1680
               Width           =   4575
            End
            Begin VB.TextBox txt_urinother 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   960
               MaxLength       =   60
               TabIndex        =   123
               Top             =   1320
               Width           =   4575
            End
            Begin VB.TextBox txt_bacteria 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   3840
               MaxLength       =   25
               TabIndex        =   121
               Top             =   600
               Width           =   1695
            End
            Begin VB.TextBox txt_casts 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   960
               MaxLength       =   15
               TabIndex        =   120
               Top             =   600
               Width           =   1455
            End
            Begin VB.TextBox txt_squamous 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   3840
               MaxLength       =   10
               TabIndex        =   122
               Top             =   960
               Width           =   1695
            End
            Begin VB.TextBox txt_wbc 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   960
               MaxLength       =   10
               TabIndex        =   118
               Text            =   "l"
               Top             =   240
               Width           =   1455
            End
            Begin VB.TextBox txt_rbc 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   3840
               MaxLength       =   10
               TabIndex        =   119
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Remarks"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   120
               TabIndex        =   157
               Top             =   1680
               Width           =   750
            End
            Begin VB.Label lbl_urinother 
               AutoSize        =   -1  'True
               Caption         =   "Others"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   300
               TabIndex        =   156
               Top             =   1320
               Width           =   570
            End
            Begin VB.Label lbl_bacteria 
               AutoSize        =   -1  'True
               Caption         =   "Bacteria"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   3030
               TabIndex        =   155
               Top             =   600
               Width           =   720
            End
            Begin VB.Label lbl_casts 
               AutoSize        =   -1  'True
               Caption         =   "Casts"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   510
               TabIndex        =   154
               Top             =   600
               Width           =   360
            End
            Begin VB.Label lbl_squamous 
               AutoSize        =   -1  'True
               Caption         =   "Squamous Epithelial Cells"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   1560
               TabIndex        =   153
               Top             =   960
               Width           =   2190
            End
            Begin VB.Label lbl_wbc 
               AutoSize        =   -1  'True
               Caption         =   "WBC"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   435
               TabIndex        =   152
               Top             =   240
               Width           =   435
            End
            Begin VB.Label lbl_rbc 
               AutoSize        =   -1  'True
               Caption         =   "RBC"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   3360
               TabIndex        =   151
               Top             =   240
               Width           =   390
            End
         End
         Begin VB.Frame frm_urine_macro 
            Caption         =   "Urine Macroscopy Examination"
            ForeColor       =   &H000000C0&
            Height          =   2055
            Left            =   120
            TabIndex        =   101
            Top             =   120
            Width           =   5535
            Begin VB.ComboBox cmb_nitrite 
               Height          =   315
               Left            =   3840
               TabIndex        =   220
               Top             =   1680
               Width           =   1455
            End
            Begin VB.TextBox txt_urinecolour 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   1080
               MaxLength       =   25
               TabIndex        =   109
               Top             =   240
               Width           =   1455
            End
            Begin VB.TextBox txt_gravity 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   3840
               MaxLength       =   5
               TabIndex        =   110
               Top             =   240
               Width           =   1455
            End
            Begin VB.ComboBox cmb_ph 
               Height          =   315
               Left            =   1080
               TabIndex        =   111
               Top             =   600
               Width           =   1455
            End
            Begin VB.ComboBox cmb_glucose 
               Height          =   315
               Left            =   1080
               TabIndex        =   113
               Top             =   960
               Width           =   1455
            End
            Begin VB.ComboBox cmb_bilirubin 
               Height          =   315
               Left            =   1080
               TabIndex        =   115
               Top             =   1320
               Width           =   1455
            End
            Begin VB.ComboBox cmb_blood 
               Height          =   315
               Left            =   1080
               TabIndex        =   117
               Top             =   1680
               Width           =   1455
            End
            Begin VB.ComboBox cmb_proteins 
               Height          =   315
               Left            =   3840
               TabIndex        =   112
               Top             =   600
               Width           =   1455
            End
            Begin VB.ComboBox cmb_ketones 
               Height          =   315
               Left            =   3840
               TabIndex        =   114
               Top             =   960
               Width           =   1455
            End
            Begin VB.ComboBox cmb_uroblngn 
               Height          =   315
               Left            =   3840
               TabIndex        =   116
               Top             =   1320
               Width           =   1455
            End
            Begin VB.TextBox txt_ph 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   1080
               MaxLength       =   10
               TabIndex        =   102
               Top             =   600
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.TextBox txt_protins 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   3840
               MaxLength       =   15
               TabIndex        =   103
               Top             =   600
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.TextBox txt_glucose 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   1080
               MaxLength       =   15
               TabIndex        =   104
               Top             =   960
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.TextBox txt_ketones 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   3840
               MaxLength       =   15
               TabIndex        =   105
               Top             =   960
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.TextBox txt_bilirubin 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   1080
               MaxLength       =   15
               TabIndex        =   106
               Top             =   1320
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.TextBox txt_uro 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   3840
               MaxLength       =   15
               TabIndex        =   107
               Top             =   1320
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.TextBox txt_blood 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   1080
               MaxLength       =   15
               TabIndex        =   108
               Top             =   1680
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Label Label14 
               Caption         =   "Nitrite"
               Height          =   255
               Left            =   2760
               TabIndex        =   219
               Top             =   1680
               Width           =   855
            End
            Begin VB.Label lbl_urinecolour 
               AutoSize        =   -1  'True
               Caption         =   "Colour"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   480
               TabIndex        =   149
               Top             =   240
               Width           =   555
            End
            Begin VB.Label lbl_gravity 
               AutoSize        =   -1  'True
               Caption         =   "Sp. Gravity"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2865
               TabIndex        =   148
               Top             =   240
               Width           =   960
            End
            Begin VB.Label lbl_ph 
               AutoSize        =   -1  'True
               Caption         =   "pH"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   780
               TabIndex        =   147
               Top             =   600
               Width           =   255
            End
            Begin VB.Label lbl_protins 
               AutoSize        =   -1  'True
               Caption         =   "Proteins"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   3120
               TabIndex        =   146
               Top             =   600
               Width           =   705
            End
            Begin VB.Label lbl_bilirubin 
               AutoSize        =   -1  'True
               Caption         =   "Bilirubin "
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   285
               TabIndex        =   145
               Top             =   1320
               Width           =   750
            End
            Begin VB.Label lbl_ketones 
               AutoSize        =   -1  'True
               Caption         =   "Ketones"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   3120
               TabIndex        =   144
               Top             =   960
               Width           =   705
            End
            Begin VB.Label lbl_glucose 
               AutoSize        =   -1  'True
               Caption         =   "Glucose"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   330
               TabIndex        =   143
               Top             =   960
               Width           =   705
            End
            Begin VB.Label lbl_uro 
               AutoSize        =   -1  'True
               Caption         =   "Urobilinogen"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2715
               TabIndex        =   142
               Top             =   1320
               Width           =   1080
            End
            Begin VB.Label lbl_blood 
               AutoSize        =   -1  'True
               Caption         =   "Blood"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   540
               TabIndex        =   141
               Top             =   1680
               Width           =   495
            End
         End
      End
      Begin VB.Frame frm_stoolexam 
         Height          =   2175
         Left            =   120
         TabIndex        =   81
         Top             =   2400
         Width           =   11535
         Begin VB.Frame frm_stool_macro 
            Caption         =   "Stool Macroscopy Examination"
            ForeColor       =   &H000000C0&
            Height          =   2055
            Left            =   120
            TabIndex        =   94
            Top             =   120
            Width           =   3015
            Begin VB.TextBox txt_consistency 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   1200
               MaxLength       =   20
               TabIndex        =   126
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox txt_stoolcolour 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   1200
               MaxLength       =   25
               TabIndex        =   125
               Top             =   240
               Width           =   1575
            End
            Begin VB.TextBox txt_parasties 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   1200
               MaxLength       =   10
               TabIndex        =   129
               Top             =   1680
               Width           =   1575
            End
            Begin VB.TextBox txt_mucus 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   1200
               MaxLength       =   10
               TabIndex        =   127
               Top             =   960
               Width           =   1575
            End
            Begin VB.TextBox txt_sblood 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   1200
               MaxLength       =   15
               TabIndex        =   128
               Top             =   1320
               Width           =   1575
            End
            Begin VB.Label lbl_parasties 
               AutoSize        =   -1  'True
               Caption         =   "Parasties"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   360
               TabIndex        =   99
               Top             =   1680
               Width           =   795
            End
            Begin VB.Label lbl_stoolblood 
               AutoSize        =   -1  'True
               Caption         =   "Blood"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   600
               TabIndex        =   98
               Top             =   1320
               Width           =   495
            End
            Begin VB.Label lbl_mucus 
               AutoSize        =   -1  'True
               Caption         =   "Mucus"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   600
               TabIndex        =   97
               Top             =   960
               Width           =   570
            End
            Begin VB.Label lbl_consistency 
               AutoSize        =   -1  'True
               Caption         =   "Consistency"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   120
               TabIndex        =   96
               Top             =   600
               Width           =   1035
            End
            Begin VB.Label lbl_stoolcolour 
               AutoSize        =   -1  'True
               Caption         =   "Colour"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   600
               TabIndex        =   95
               Top             =   240
               Width           =   555
            End
         End
         Begin VB.Frame frm_stool_micro 
            Caption         =   "Stool Microscopy Examination"
            ForeColor       =   &H000000C0&
            Height          =   2055
            Left            =   3240
            TabIndex        =   82
            Top             =   120
            Width           =   8055
            Begin VB.TextBox txt_stoolcomments 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   960
               MaxLength       =   100
               TabIndex        =   140
               Top             =   1680
               Width           =   6975
            End
            Begin VB.TextBox txt_stoolother 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   960
               MaxLength       =   200
               TabIndex        =   139
               Top             =   1320
               Width           =   6975
            End
            Begin VB.TextBox txt_occult 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   6360
               MaxLength       =   15
               TabIndex        =   138
               Top             =   960
               Width           =   1575
            End
            Begin VB.TextBox txt_fat 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   3360
               MaxLength       =   6
               TabIndex        =   137
               Top             =   960
               Width           =   1575
            End
            Begin VB.TextBox txt_starch 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   960
               MaxLength       =   6
               TabIndex        =   136
               Top             =   960
               Width           =   1575
            End
            Begin VB.TextBox txt_microphages 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   6360
               MaxLength       =   6
               TabIndex        =   135
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox txt_stoolrbc 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   3360
               MaxLength       =   10
               TabIndex        =   134
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox txt_puscells 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   960
               MaxLength       =   10
               TabIndex        =   133
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox txt_larvae 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   6360
               MaxLength       =   100
               TabIndex        =   132
               Top             =   240
               Width           =   1575
            End
            Begin VB.TextBox txt_Ova 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   960
               MaxLength       =   100
               TabIndex        =   130
               Top             =   240
               Width           =   1575
            End
            Begin VB.TextBox txt_cysts 
               Appearance      =   0  'Flat
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   3360
               MaxLength       =   100
               TabIndex        =   131
               Top             =   240
               Width           =   1575
            End
            Begin VB.ComboBox txt_stoolcomments1 
               Height          =   315
               Left            =   960
               TabIndex        =   213
               Text            =   "Combo1"
               Top             =   1680
               Visible         =   0   'False
               Width           =   6975
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Remarks"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   90
               TabIndex        =   93
               Top             =   1680
               Width           =   750
            End
            Begin VB.Label lbl_stoolother 
               AutoSize        =   -1  'True
               Caption         =   "Others"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   270
               TabIndex        =   92
               Top             =   1320
               Width           =   570
            End
            Begin VB.Label lbl_occult 
               AutoSize        =   -1  'True
               Caption         =   "Occult blood"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   5145
               TabIndex        =   91
               Top             =   960
               Width           =   1095
            End
            Begin VB.Label lbl_fat 
               AutoSize        =   -1  'True
               Caption         =   "Fat "
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2880
               TabIndex        =   90
               Top             =   960
               Width           =   345
            End
            Begin VB.Label lbl_starch 
               AutoSize        =   -1  'True
               Caption         =   "Starch"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   270
               TabIndex        =   89
               Top             =   960
               Width           =   570
            End
            Begin VB.Label lbl_macrophages 
               AutoSize        =   -1  'True
               Caption         =   "Macrophages"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   5085
               TabIndex        =   88
               Top             =   600
               Width           =   1155
            End
            Begin VB.Label lbl_stoolrbc 
               AutoSize        =   -1  'True
               Caption         =   "RBC"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2835
               TabIndex        =   87
               Top             =   600
               Width           =   390
            End
            Begin VB.Label lbl_puscells 
               AutoSize        =   -1  'True
               Caption         =   "Pus cells"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   60
               TabIndex        =   86
               Top             =   600
               Width           =   780
            End
            Begin VB.Label lbl_larvae 
               AutoSize        =   -1  'True
               Caption         =   "Larvae"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   5640
               TabIndex        =   85
               Top             =   240
               Width           =   600
            End
            Begin VB.Label lbl_cysts 
               AutoSize        =   -1  'True
               Caption         =   "Cysts"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2760
               TabIndex        =   84
               Top             =   240
               Width           =   465
            End
            Begin VB.Label lbl_Ova 
               AutoSize        =   -1  'True
               Caption         =   "Ova"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   480
               TabIndex        =   83
               Top             =   240
               Width           =   360
            End
         End
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4575
      Left            =   120
      TabIndex        =   68
      Top             =   3240
      Visible         =   0   'False
      Width           =   11775
      Begin VB.ComboBox cmbmicrodrug 
         Height          =   315
         Left            =   3840
         TabIndex        =   247
         Text            =   "Select"
         Top             =   3240
         Width           =   1815
      End
      Begin VB.TextBox text1 
         Height          =   285
         Left            =   2520
         TabIndex        =   71
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox txt_MCcomment 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   224
         Top             =   3840
         Width           =   6735
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2415
         Left            =   2520
         TabIndex        =   72
         Top             =   1200
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   4260
         _Version        =   393216
         Cols            =   4
         BackColor       =   16777215
         AllowBigSelection=   -1  'True
      End
      Begin VB.TextBox txt_test 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         TabIndex        =   70
         Top             =   240
         Width           =   4095
      End
      Begin VB.TextBox txt_testresult 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         TabIndex        =   69
         Top             =   720
         Width           =   4095
      End
      Begin VB.Label Label16 
         Caption         =   "Comment"
         Height          =   255
         Left            =   1440
         TabIndex        =   223
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label10 
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
         Left            =   2520
         TabIndex        =   74
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label9 
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
         Left            =   2520
         TabIndex        =   73
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.Frame frmSerology 
      Height          =   4695
      Left            =   0
      TabIndex        =   228
      Top             =   3120
      Visible         =   0   'False
      Width           =   11895
      Begin VB.TextBox txtLimitation 
         Height          =   495
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   246
         Top             =   3720
         Width           =   9975
      End
      Begin VB.ComboBox cmbProCalCon 
         Height          =   315
         Left            =   1680
         TabIndex        =   244
         Top             =   1680
         Width           =   6255
      End
      Begin VB.ComboBox cmbProcalci 
         Height          =   315
         Left            =   1680
         TabIndex        =   243
         Top             =   1320
         Width           =   3735
      End
      Begin VB.ComboBox cmbSerRange 
         Height          =   315
         Left            =   1680
         TabIndex        =   242
         Top             =   960
         Width           =   6255
      End
      Begin VB.ComboBox cmbSerResult 
         Height          =   315
         Left            =   1680
         TabIndex        =   241
         Top             =   600
         Width           =   9975
      End
      Begin VB.TextBox txtSerTestPerform 
         Height          =   285
         Left            =   1680
         TabIndex        =   240
         Top             =   240
         Width           =   9975
      End
      Begin VB.TextBox txtInBold 
         Height          =   285
         Left            =   1680
         TabIndex        =   236
         Top             =   4320
         Width           =   9015
      End
      Begin VB.TextBox txtMethod 
         Height          =   615
         Left            =   1680
         TabIndex        =   234
         Top             =   3000
         Width           =   9975
      End
      Begin VB.TextBox txtInterpretation 
         Height          =   855
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   232
         Top             =   2040
         Width           =   9975
      End
      Begin VB.Label Label26 
         Caption         =   "Limitation"
         Height          =   255
         Left            =   120
         TabIndex        =   245
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "Procalcitonin concentration"
         Height          =   375
         Left            =   120
         TabIndex        =   239
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label24 
         Caption         =   "Procalcitonin"
         Height          =   255
         Left            =   120
         TabIndex        =   238
         Top             =   1370
         Width           =   1215
      End
      Begin VB.Label Label23 
         Caption         =   "Range of Test"
         Height          =   255
         Left            =   120
         TabIndex        =   237
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label22 
         Caption         =   "Text in Bold"
         Height          =   255
         Left            =   120
         TabIndex        =   235
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label Label21 
         Caption         =   "Method"
         Height          =   255
         Left            =   120
         TabIndex        =   233
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label20 
         Caption         =   "Interpretation"
         Height          =   255
         Left            =   120
         TabIndex        =   231
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label19 
         Caption         =   "Result"
         Height          =   255
         Left            =   120
         TabIndex        =   230
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label18 
         Caption         =   "Test Performed"
         Height          =   255
         Left            =   120
         TabIndex        =   229
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Label Label27 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1920
      TabIndex        =   250
      Top             =   9120
      Width           =   8415
   End
   Begin VB.Label lbl_microbiology 
      AutoSize        =   -1  'True
      Caption         =   "-REPORT - MICROBIOLOGY "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   3840
      TabIndex        =   208
      Top             =   360
      Width           =   3990
   End
   Begin VB.Label lblicu 
      Caption         =   "ICU Patient"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   0
      TabIndex        =   207
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "TATA MEMORIAL HOSPITAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3000
      TabIndex        =   206
      Top             =   0
      Width           =   5775
   End
   Begin VB.Label Label8 
      Caption         =   "14/12/10"
      Height          =   255
      Left            =   0
      TabIndex        =   205
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "Report_microbiology1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public a As Integer
Public req
Option Explicit
''SHWETA 18-OCT-2010
Dim DBAANTI_ARR(1000) As String
''''''''''''''''''''''
Dim callstatus   As Boolean
Dim srno  As Integer
'***** 23/10/08
Dim q As Boolean

Dim RQN As String
Dim str As String
Dim countr As Integer

Dim GINS_QR As rdoQuery
Dim GUPD_QR As rdoQuery

Dim ins_mccomment_qr As rdoQuery

Dim GBINS_QR As rdoQuery
Dim GBUPD_QR As rdoQuery

Dim UPDReqRQ As rdoQuery
Dim UPDReqBRQ As rdoQuery 'Shraddha Pandit 14/11/08

Dim upd_mccomment_qr As rdoQuery

Dim sel_comment_qr As rdoQuery
Dim sel_comment_rs As rdoResultset


Dim ins_mod_var As String 'Shraddha Pandit 05/06/09

Dim sel_dreqpc_qr As rdoQuery 'Shraddha Pandit 17/02/10

Dim sel_dreqpc_rs As rdoResultset 'Shraddha Pandit 17/02/10

'Dim UPDREQRS As rdoResultset

Dim ViewMTQR As rdoQuery
Dim ViewMTRS As rdoResultset


Dim ViewGBQR As rdoQuery
Dim ViewGBRS As rdoResultset

Dim SELDRUG_QR As rdoQuery
Dim SELDRUG_RS As rdoResultset

'**************
Dim VAR_REQNNO As String
Dim objWord             As Word.Application
Dim objDoc              As Word.Document
'for selecting antibiotics
Dim sel_antiqr          As rdo.rdoQuery
Dim sel_moadlityqr      As rdo.rdoQuery
Dim sel_deptqr          As rdo.rdoQuery
Dim req_qr              As rdo.rdoQuery
Dim SEL_detail_qr       As rdo.rdoQuery ' for selecting details of patient
Dim SELRREQUISITN_QR    As rdo.rdoQuery
Dim sel_proccode_qr     As rdo.rdoQuery ' for selecting procedure code
Dim sel_submodali_qr    As rdo.rdoQuery ' for selecting submodality
Dim REFDR_QR            As rdo.rdoQuery ' for refered by doctor not in used
Dim sel_bed_qr          As rdo.rdoQuery ' to select bed and ward
Dim SEL_MODALITY_QR     As rdo.rdoQuery
Dim sel_forupd2_qr      As rdo.rdoQuery ' for antibiotic
Dim sel_forupd1_qr      As rdo.rdoQuery ' toget data from table for bactriology
Dim sel_forupd3_qr      As rdo.rdoQuery ' to get data from table for urine
Dim sel_forupd4_qr      As rdo.rdoQuery ' to get data from table for stool
Dim sel_forupd5_qr      As rdo.rdoQuery ' to get data from table for other
Dim sel_forupd6_qr      As rdo.rdoQuery ' to get data from mycobacteriology table
Dim sel_micro_qr        As rdo.rdoQuery ' to  get micro number
Dim sel_ris_qr          As rdo.rdoQuery ' to get status of commit flag
Dim sel_FINAL_qr        As rdo.rdoQuery ' to get status of commit flag
Dim sel_ENTER_qr        As rdo.rdoQuery ' to get ENTER BY USER LIST
Dim SEL_dreq_staff_qr   As rdo.rdoQuery ' to get staff detail
Dim upd_12_qr           As rdo.rdoQuery ' to update record in dmbaanti table
Dim ins_0_qr            As rdo.rdoQuery ' insert query
Dim ins_1_qr            As rdo.rdoQuery
Dim ins_2_qr            As rdo.rdoQuery
Dim insGZZqr As rdo.rdoQuery 'vidya on 05/05/2012
Dim UpdGZZqr As rdo.rdoQuery 'vidya on 05/05/2012
Dim ins_4_qr            As rdo.rdoQuery
Dim ins_5_qr            As rdo.rdoQuery
Dim ins_6_qr            As rdo.rdoQuery 'insert in mycobacteriology table
'Dim ins_7_qr            As RDO.rdoQuery 'insert in DMURE TABLE AS DUMMY REQNO
'Dim ins_8_qr            As RDO.rdoQuery 'insert in DMSRE TABLE AS DUMMY REQNO
Dim upd_0_qr            As rdo.rdoQuery 'update query
Dim upd_1_qr            As rdo.rdoQuery
Dim upd_2_qr            As rdo.rdoQuery
Dim upd_4_qr            As rdo.rdoQuery
Dim upd_5_qr            As rdo.rdoQuery
Dim upd_6_qr            As rdo.rdoQuery
Dim upd_7_qr            As rdo.rdoQuery 'for updating mycobacteriology
Dim upd_8_qr            As rdo.rdoQuery 'for updating rislib about enterby
Dim upd_11_qr           As rdo.rdoQuery 'for setting requistion status
Dim del_dmbaanti_qr     As rdo.rdoQuery 'for deleting from dmbaanti table if pathogen is not grown
Dim del_dmbaanti_qr1    As rdo.rdoQuery
Dim I                   As Integer      'USE FOR SELECTING INDEX OF MICROBIOLOGY in combo
Dim sel_antirs          As rdo.rdoResultset
Dim sel_modalityrs      As rdo.rdoResultset
Dim sel_deptrs          As rdo.rdoResultset
Dim SEL_MODALITY_RS     As rdo.rdoResultset
Dim req_rs              As rdo.rdoResultset
Dim SEL_detail_rs       As rdo.rdoResultset
'******** Shraddha 26/11/08
Dim SELSPONANTIBRS As rdoResultset
Dim SELSPONANTIBQR As rdoQuery
'********

'************ Shraddha Pandit 23/01/09 ****************
Dim SEL_DMUL_QR As rdoQuery
Dim SEL_DMUL_RS As rdoResultset
'******************************************************

Dim sel_BAcomment_qr As rdoQuery
Dim sel_BAcomment_rs As rdoResultset


Dim SELRREQUISITN_RS    As rdo.rdoResultset
Dim sel_proccode_rs     As rdo.rdoResultset 'for selecting procedure code
Dim sel_submodali_rs    As rdo.rdoResultset 'for selecting submodality
Dim REFDR_RS            As rdo.rdoResultset 'not in used
Dim sel_bed_rs          As rdo.rdoResultset
Dim sel_forupd1_rs      As rdo.rdoResultset
Dim sel_forupd2_rs      As rdo.rdoResultset
Dim sel_forupd3_rs      As rdo.rdoResultset
Dim sel_forupd4_rs      As rdo.rdoResultset
Dim sel_forupd5_rs      As rdo.rdoResultset
Dim sel_forupd6_rs      As rdo.rdoResultset
Dim sel_micro_rs        As rdo.rdoResultset ' to  get micro number
Dim sel_ris_rs          As rdo.rdoResultset ' to  get micro number
Dim sel_ENTER_rs        As rdo.rdoResultset ' to  get ENTER BY USER LIST
Dim sel_FINAL_rs        As rdo.rdoResultset ' to  get FINAL BY USER LIST
Dim SEL_dreq_staff_rs   As rdo.rdoResultset ' to get staff detail
Dim S1                  As String ' FOR DEPTCODE
Dim S2                  As String ' FOR MODALITYCODE
Dim mod_reqno           As String 'used to store reqno by calling req_chglen method
Dim new_flag            As Boolean
Dim mod_flag            As Boolean
Public VIEW_FLAG           As Boolean
Dim no_req              As Boolean 'to check before printing report whether req is existing or not
'here
''*******ICU**********
Dim SELICU_QR As rdoQuery
Dim SELICU_RS As rdoResultset
Dim SELICU_FLAG As Boolean

Dim SELMEMO_QR As rdoQuery
Dim SELMEMO_RS As rdoResultset
Dim SELMEMO_FLAG As Boolean

''*******ICU**********


Dim GZZREQPATH_QR As rdoQuery
Dim GZZREQPATH_RS As rdoResultset

Dim sel_MOtemplet_qr As rdoQuery
Dim sel_MOtemplet_rs As rdoResultset

'for addendum
Dim INSDADDENDUM_QR As rdoQuery
Dim SELDADDENDUM_QR As rdoQuery
Dim SELDADDENDUM_RS As rdoResultset
Dim SELDADDENDUMFLAG As Boolean
Dim ADDENDUMFALG As Boolean
Dim w_grossdes As Word.Application
Dim d_grossdesc As Word.Document
Dim UPDRTIMEC_QR As rdoQuery
Dim UPDRTIMEC_QR1 As rdoQuery
Dim UPDRTIMEC_QR2 As rdoQuery
'''''''''''

' FOR CHECKING RECORDS PRESENCE DURING MODIFY
''********************************
Dim sel_formod2_qr As rdoQuery
Dim sel_formod2_rs As rdoResultset
Dim sel_formod2_flg As Boolean
''********************************

''******SECURITY***********

Dim LOGINUSRNAME As String
Dim LOGINTYP As String
Dim LOGINUSRID As String
Dim REPWTRNRS As rdoResultset
Dim REPWTRIRS As rdoResultset
Dim CURUSRDET_QR As rdoQuery
Dim CURUSRDET_RS As rdoResultset

Dim SELFINALBY_QR As rdoQuery
Dim SELFINALBY_RS As rdoResultset
Dim REPWTRN_QR As rdoQuery
Dim txt_c_modflg As Boolean
''******SECURITY***********

'''DECOMMIT
Dim UPDRREQ_QR As rdoQuery
Dim SELDMBSEMCOM_QR As rdoQuery
Dim SELDMBSEMCOM_RS As rdoResultset
Dim SELDMBFLAG As Boolean
Dim CHECKREC_QR As rdoQuery
Dim CHECKREC_RS As rdoResultset
Dim CHECKFLAG As Boolean

''KANCHAN

Dim retantibodies_qr As rdoQuery
Dim RETRESULT_RS As rdoResultset
Dim REPORT_QR As rdoQuery
Dim REPORT_RS As rdoResultset
Dim REPORTSTR As String
Dim REPORTFLAG As Boolean
Dim UPDATERTIMELOG As rdoQuery
Dim UPDATERTIMELOG1 As rdoQuery
Dim INSINDMBANTI_QR As rdoQuery
Dim COL1 As Integer
Dim row1 As Integer
Dim MODE As String
Dim GRIDSCROLL As Boolean

Dim DEPT_CODE As String
Dim MOD_CODE As String
Dim modtxt As String
Dim ind As Integer
Dim NEWF As Boolean

Dim prev_flag As Boolean 'Added by Shraddha on 11/06/09
Dim sel_moadlityqr1 As rdoQuery 'Shraddha 16/06/2009
Dim sel_modalityrs1 As rdoResultset 'Shraddha 16/06/2009

Dim upd_rtime_qr As rdoQuery 'Shraddha 19/03/11
Dim sel_rtimelog_qr As rdoQuery 'Shraddha 22/03/11
Dim sel_rtimelog_rs As rdoResultset 'Shraddha 22/03/11

Dim PRO_CODE As String  'vidya 05/05/2012
Dim RS As rdoResultset
Dim w As Word.Application
Dim d As Word.Document

Private Sub cmb_addendumby_click()
    cmd_save.Enabled = True
    ADDENDUMFALG = True
End Sub

Private Sub cmb_antibiotic_Change()
'grid_dmbanti.TextMatrix(grid_dmbanti.row, 0) = Trim(cmb_antibiotic.Text)
End Sub

Private Sub cmb_antibiotic_Click()
'ind = 0
'SHWETA 18-OCT-2010
Dim CHECKDUP As Boolean
If grid_dmbanti.row = 1 Then
CHECKDUP = False
Else
CHECKDUP = CHECKDUPLICATE(Trim(cmb_antibiotic.Text))
End If

If CHECKDUP = True Then
cmb_antibiotic.Text = ""
Else
''*****TILL HERE
grid_dmbanti.TextMatrix(grid_dmbanti.row, 0) = Trim(cmb_antibiotic.Text)
''*****
DBAANTI_ARR(grid_dmbanti.row) = Trim(cmb_antibiotic.Text)
End If
''*****TILL HERE
'            ind = cmb_antibiotic.ListIndex
'            cmb_antibiotic.RemoveItem (ind) 'Commented on 04/06/09 Shraddhs
  
End Sub
Function CHECKDUPLICATE(ByVal ANTI As String) As Boolean
CHECKDUPLICATE = False
Dim DBI As Integer
'For DBI = 0 To (UBound(DBAANTI_ARR) - 1)
For DBI = 1 To grid_dmbanti.Rows - 1
    'If UCase(Trim(ANTI)) = UCase(Trim(DBAANTI_ARR(DBI))) Then
    'Commented and added by Shraddha on 14/12/10
    If UCase(Trim(ANTI)) = UCase(Trim(grid_dmbanti.TextMatrix(DBI, 0))) Then
    CHECKDUPLICATE = True
    MsgBox UCase(Trim(ANTI)) & " Already entered"
    Exit Function
    End If
Next
End Function

Function CHECKDUPLICATE_MB(ByVal ANTI As String) As Boolean
CHECKDUPLICATE_MB = False
Dim DBI As Integer

For DBI = 1 To grid_c_anti.Rows - 1
    'If UCase(Trim(ANTI)) = UCase(Trim(DBAANTI_ARR(DBI))) Then
    'Commented and added by Shraddha on 14/12/10
    If UCase(Trim(ANTI)) = UCase(Trim(grid_c_anti.TextMatrix(DBI, 0))) Then
    CHECKDUPLICATE_MB = True
    MsgBox UCase(Trim(ANTI)) & " Already entered"
    Exit Function
    End If
    
Next
End Function


Function IsRepeated(antinm As String) As Boolean
    IsRepeated = True
    Dim I As Integer
    With grid_dmbanti
        If .Rows <= 2 Then
            IsRepeated = False
            Exit Function
        Else
            For I = 1 To .Rows - 2
                If .TextMatrix(I, 0) = Trim(antinm) And .row <> I Then
                    IsRepeated = True
                    MsgBox "This antibiotic is already entered", vbCritical
                    
                    Exit Function
                Else
                    IsRepeated = False
                End If
            Next I
        End If
    End With
End Function

Private Sub remoanti()
Dim rowcn As Integer
Dim cn As Integer
rowcn = grid_dmbanti.Rows - 1

For cn = 1 To rowcn
    ind = 0
    cmb_antibiotic.Enabled = True
    cmb_antibiotic.Text = Trim(grid_dmbanti.TextMatrix(cn, 0))
    ind = cmb_antibiotic.ListIndex
   ' cmb_antibiotic.RemoveItem (ind)
Next cn

End Sub


Private Sub cmb_antibiotic_KeyPress(KeyAscii As Integer)

'ind = 0

If KeyAscii = 13 Then
    If cmb_antibiotic.Text = "" Then
        MsgBox "Select Antibiotic", vbOKOnly
        
        cmb_antibiotic.SetFocus
    Else
         
            Call movecontrols(, 1)
            Call GridControls("VISIBLE", 1)
                 
 '           ind = cmb_antibiotic.ListIndex
 '           cmb_antibiotic.RemoveItem (ind)
       
    End If
End If
End Sub
Public Sub MYDELAY(inttime As String)

Dim sttime
Dim endtime1 As Long
Dim TESTTIME As Long

'sttime = Timer

endtime1 = Timer + inttime
TESTTIME = Timer

Do While Timer < endtime1
 '   Text1 = endtime1
 '   Text2 = TESTTIME
    DoEvents
  '  TESTTIME = Timer
Loop


End Sub


Private Sub cmb_antibiotic_LostFocus()

If Len(Trim(cmb_antibiotic.Text)) > 0 Then
ind = 0
'
ind = cmb_antibiotic.ListIndex

If ind = -1 Then
Else
'cmb_antibiotic.RemoveItem (ind) 'Commented on 05/06/09 by Shraddha
End If
End If
'GET_ANTIBIOTIC
'If IsRepeated(Trim(cmb_antibiotic.Text)) Then
'
'            cmb_antibiotic.Text = ""
'            cmb_antibiotic.SetFocus
'  Else
'            grid_dmbanti.TextMatrix(grid_dmbanti.row, 0) = Trim(cmb_antibiotic.Text)
'  End If
End Sub


Private Sub cmb_antiboitic_mb_Click()
Dim CHECKDUP As Boolean
If grid_c_anti.row = 1 Then
    CHECKDUP = False
Else
    CHECKDUP = CHECKDUPLICATE_MB(Trim(cmb_antiboitic_mb.Text))
End If

If CHECKDUP = True Then
    cmb_antiboitic_mb.Text = ""
Else

grid_c_anti.TextMatrix(grid_c_anti.row, 0) = Trim(cmb_antiboitic_mb.Text)

DBAANTI_ARR(grid_c_anti.row) = Trim(cmb_antiboitic_mb.Text)
End If
''*****TILL HERE
'            ind = cmb_antibiotic.ListIndex
'            cmb_antibiotic.RemoveItem (ind) 'Commented on 04/06/09 Shraddhs
End Sub

Private Sub cmb_antiboitic_mb_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cmb_antiboitic_mb.Text = "" Then
        MsgBox "Select Antibiotic", vbOKOnly
        cmb_antiboitic_mb.SetFocus
    Else
        Call movecontrols_MB(, 1)
        Call GridControls_mb("VISIBLE", 1)
    End If
End If
End Sub

Private Sub cmb_antiboitic_mb_LostFocus()
If Len(Trim(cmb_antiboitic_mb.Text)) > 0 Then
    ind = 0
    ind = cmb_antiboitic_mb.ListIndex
    If ind = -1 Then
    
    Else
        'cmb_antibiotic.RemoveItem (ind) 'Commented on 05/06/09 by Shraddha
    End If
    txt_c_1.Visible = True
    txt_c_1.SetFocus
End If
'GET_ANTIBIOTIC
'If IsRepeated(Trim(cmb_antibiotic.Text)) Then
'
'            cmb_antibiotic.Text = ""
'            cmb_antibiotic.SetFocus
'  Else
'            grid_dmbanti.TextMatrix(grid_dmbanti.row, 0) = Trim(cmb_antibiotic.Text)
'  End If
End Sub

Private Sub cmb_modality_Click()
On Error GoTo MSG
    I = InStr(cmb_modality.Text, "$")
    txt_modcode.Text = Mid(cmb_modality.Text, I + 1, 2)
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub cmb_modality_LostFocus()
 If cmb_modality.Text = "" Then
        MsgBox "Select Modality First", vbOKOnly
    End If
End Sub



Private Sub cmb_motemplet_Click()
Dim MICROMO_QR As rdoQuery
Dim MICROMO_RS As rdoResultset

txt_test_perform.Text = Trim(cmb_motemplet.Text)

str = "SELECT * FROM DISLIB.MICROMO WHERE TESTNAME=?"
Set MICROMO_QR = rdb.CreateQuery("", str)

MICROMO_QR(0) = Trim(cmb_motemplet.Text)
Set MICROMO_RS = MICROMO_QR.OpenResultset(1, 3)

If Not MICROMO_RS.EOF Then
    txt_result.Text = Trim(MICROMO_RS("DESCRIPN")) & vbCrLf & Trim(MICROMO_RS("FLD1")) & vbCrLf & Trim(MICROMO_RS("FLD2")) & vbCrLf & Trim(MICROMO_RS("FLD3")) & vbCrLf & Trim(MICROMO_RS("FLD4")) & vbCrLf & Trim(MICROMO_RS("FLD5"))
End If
End Sub

Private Sub cmb_org1_LostFocus()
'******** Shraddha 03/07/09 ********
Dim o1, m1
Dim N1, L1 As Integer
L1 = Len(cmb_org1.Text)
N1 = InStr(1, cmb_org1.Text, "$")

If N1 <> 0 Then
    m1 = Mid(cmb_org1, N1 + 1, L1 - N1)
    o1 = Mid(cmb_org1, 1, N1 - 1)
    
    If m1 = "Heavy" Then
        cmb_org1.Text = "Heavy growth of " & o1
    End If
    
    If m1 = "Poor" Then
        cmb_org1.Text = "Poor growth of " & o1
    End If
    
    If m1 = "Moderate" Then
        cmb_org1.Text = "Moderate growth of " & o1
    End If
    
    If m1 = "Significant" Then
        cmb_org1.Text = "Significant growth of " & o1
    End If
    'cmb_org2.SetFocus 'modified by mrunal 26_2
End If
'SendKeys "^{f4}"

'************************************
End Sub

    Private Sub cmb_org2_LostFocus()
'******** Shraddha 03/07/09 ********

Dim o2, m2
Dim n2, l2 As Integer
l2 = Len(cmb_org2.Text)
n2 = InStr(1, cmb_org2.Text, "$")

If n2 <> 0 Then

    If l2 > n2 Then  'Shraddha 25/11/10

    m2 = Mid(cmb_org2, n2 + 1, l2 - n2)
    o2 = Mid(cmb_org2, 1, n2 - 1)
    
      End If
    
          If m2 = "Heavy" Then
        cmb_org2.Text = "Heavy growth of " & o2
    End If
    
    If m2 = "Poor" Then
        cmb_org2.Text = "Poor growth of " & o2
    End If
    
    If m2 = "Moderate" Then
        cmb_org2.Text = "Moderate growth of " & o2
    End If
    
    If m2 = "Significant" Then
        cmb_org2.Text = "Significant growth of " & o2
    End If
    'cmb_org3.SetFocus 'modified by mrunal 26_2
End If
'SendKeys "^{f4}"



'************************************
End Sub

Private Sub cmb_organisms_LOSTFOCUS()
'On Error GoTo MSG
'Commented on 12/05/11 by Shraddha
'    If new_flag Then
'        sel_antiqr(0) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
'        sel_antiqr(1) = Trim(cmb_organisms.Text)
'        Set sel_antirs = sel_antiqr.OpenResultset(1, 3)
'        If Not sel_antirs.EOF Then
'            grid_c_anti.clear
'            grid_c_anti.Rows = 1
'            grid_c_anti.Cols = 2
'            grid_c_anti.FormatString = "   ANTIMICROBIAL                               |^         1           "
'            While Not sel_antirs.EOF
'              grid_c_anti.AddItem sel_antirs("antibdesc")
'              sel_antirs.MoveNext
'            Wend
'            grid_c_anti.row = 1
'            grid_c_anti.col = 1
'            txt_c_1.Width = grid_c_anti.CellWidth
'            txt_c_1.Height = grid_c_anti.CellHeight
'            txt_c_1.Left = grid_c_anti.CellLeft + grid_c_anti.Left
'            txt_c_1.Top = grid_c_anti.CellTop + grid_c_anti.Top
'            txt_c_1.Visible = True
'            sel_antirs.Close
'        Else
'            MsgBox "Use Master Screen to Enter Antimicrobial", vbOKOnly
'            sel_antirs.Close
'            Exit Sub
'        End If
'     ElseIf mod_flag Then
'        If Not sel_forupd2_rs.EOF Then
'            grid_c_anti.Rows = 1
'            grid_c_anti.FormatString = "   ANTIMICROBIAL                               |^         1           "
'            While Not sel_forupd2_rs.EOF
'                grid_c_anti.AddItem sel_forupd2_rs("antibiotic") & vbTab & sel_forupd2_rs("antibio1")
'                sel_forupd2_rs.MoveNext
'            Wend
'            sel_forupd2_rs.Close
'         Else
'            sel_antiqr(0) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
'            sel_antiqr(1) = Trim(cmb_organisms.Text)
'            Set sel_antirs = sel_antiqr.OpenResultset(1, 3)
'            If Not sel_antirs.EOF Then
'                grid_c_anti.Rows = 1
'                grid_c_anti.Cols = 2
'                grid_c_anti.FormatString = "   ANTIMICROBIAL                               |^         1           "
'                While Not sel_antirs.EOF
'                    grid_c_anti.AddItem sel_antirs("antibdesc")
'                    sel_antirs.MoveNext
'                Wend
'                grid_c_anti.row = 1
'                grid_c_anti.col = 1
'                txt_c_1.Width = grid_c_anti.CellWidth
'                txt_c_1.Height = grid_c_anti.CellHeight
'                txt_c_1.Left = grid_c_anti.CellLeft + grid_c_anti.Left
'                txt_c_1.Top = grid_c_anti.CellTop + grid_c_anti.Top
'                txt_c_1.Visible = True
'                sel_antirs.Close
'            End If
'            sel_forupd2_rs.Close
'         End If
'     End If
'     Exit Sub
' Till here 12/05/11 Shraddha
    '***** Shraddha 10/05/11 *****
    GET_ANTIBIOTIC_MB
    
    If new_flag Then
        grid_c_anti.Rows = 2
        grid_c_anti.Cols = 2
        grid_c_anti.FormatString = "   ANTIMICROBIAL                               |^         1           "
      
      Call movecontrols_MB(, 0)
      Call GridControls_mb("VISIBLE", 0)
    
    ElseIf mod_flag Then
        If Len(Trim(cmb_organisms.Text)) >= 1 Then
            
            
                    grid_c_anti.Rows = 1
                    grid_c_anti.Cols = 2
                    grid_c_anti.FormatString = "   ANTIMICROBIAL                               |^         1           "
               
            
            
        End If
        ''getresult_bacteriology
        sel_forupd2_qr(0) = mod_reqno
        Set sel_forupd2_rs = sel_forupd2_qr.OpenResultset(1, 3)
            
        If sel_forupd2_rs.RowCount > 0 Then
            While Not sel_forupd2_rs.EOF
                If grid_c_anti.Cols = 2 Then
                    While Not sel_forupd2_rs.EOF
                        grid_c_anti.AddItem sel_forupd2_rs("antibiotic") & vbTab & sel_forupd2_rs("antibio1")
                        sel_forupd2_rs.MoveNext
                    Wend
                End If
            Wend
        Else
            grid_c_anti.Rows = grid_c_anti.Rows + 1
            Call movecontrols_MB(, 0)
            Call GridControls_mb("VISIBLE", 0)
        End If
    End If
    
    'cmb_antiboitic_mb.SetFocus
    Exit Sub
    '********** 10/05/11 *********
MSG:
    MsgBox Err.Description
End Sub

Private Sub cmb_patho_click()
On Error GoTo MSG
'If cmb_patho.ItemData(cmb_patho.ListIndex) = 0 Then 'selected yes then
If cmb_patho.Text = "Yes" Then
    lbl_report.Visible = False
    txt_report.Visible = False
    lbl_organisms.Visible = True
    
    '***** Shraddha 30/06/09 *****
    'txt_organ1.Visible = True
    'TXT_ORGAN2.Visible = True
    'TXT_ORGAN3.Visible = True
    
    cmb_org1.Visible = True
    cmb_org2.Visible = True
    cmb_org3.Visible = True
    '*****************************
    
    lbl_1.Visible = True
    lbl_2.Visible = True
    lbl_3.Visible = True
    'grid_anti.Visible = True
    grid_dmbanti.Visible = True
    lbl_bcomments.Left = 600
    'lbl_bcomments.Top = 3720
    lbl_bcomments.Top = 4440 'Shraddha 03/07/09
    txt_comment.Left = 1560
    'txt_comment.Top = 3720
    txt_comment.Top = 4440 'Shraddha 03/07/09
Else 'if selected no then
    lbl_report.Visible = True
    txt_report.Visible = True
    lbl_organisms.Visible = False
    
    '***** Shraddha 30/06/09 *****
    txt_organ1.Visible = False
    TXT_ORGAN2.Visible = False
    TXT_ORGAN3.Visible = False
    
    cmb_org1.Visible = False
    cmb_org2.Visible = False
    cmb_org3.Visible = False
    '*****************************
    txt_1.Visible = False
    txt_2.Visible = False
    txt_3.Visible = False
    lbl_1.Visible = False
    lbl_2.Visible = False
    lbl_3.Visible = False
    'grid_anti.Visible = False
    grid_dmbanti.Visible = False
    lbl_report.Left = 960
    lbl_report.Top = 2040
    'txt_report.Left = 1560
    txt_report.Left = 1700 'Shraddha 03/07/09
    'txt_report.Top = 1920
    txt_report.Top = 2040 'Shraddha 03/07/09
    txt_comment.Left = 1560
    'txt_comment.Top = 2640
    txt_comment.Top = 3360 'Shraddha 03/07/09
    lbl_bcomments.Left = 600
    'lbl_bcomments.Top = 2640
    lbl_bcomments.Top = 3360 'Shraddha 03/07/09
    If mod_flag Then
        'clear previous data of dmbaanti table
        
        '****** commented and added on 30/06/09 by Shraddha ********
        'txt_organ1.Text = ""
        'TXT_ORGAN2.Text = ""
        'TXT_ORGAN3.Text = ""
        
        cmb_org1.Text = ""
        cmb_org2.Text = ""
        cmb_org3.Text = ""
        '***********************************************************
        
        del_dmbaanti_qr(0) = mod_reqno
        del_dmbaanti_qr.Execute
        'grid_anti.Rows = 1
        'grid_anti.Cols = 3
        grid_dmbanti.Rows = 1
        grid_dmbanti.Cols = 3
    End If
End If
Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub cmb_patho_LostFocus()
'******** Shraddha 03/07/09 ********
If cmb_org1.Visible Then
SendKeys "^{f4}"
cmb_org1.SetFocus
End If
'************************************
End Sub

Private Sub CMB_PROC_Change()
'If CMB_PROC.Text = "G256" Then
'Dim aa As Frame
'For Each aa In Controls
'        If TypeOf c Is CheckBox Then
'            c.Enabled = True
'        End If
'    Next
'
'Frame3.Visible = True
'End If
End Sub

Public Sub cmb_reqno_Click()
Dim gk As Integer
txt_modcode.Text = Mid(cmb_reqno.Text, 5, 2)
txt_deptcode.Text = Left(cmb_reqno.Text, 3)
''***** Shraddha 22/03/11 *****
' Dim rqno As String
' Dim j As Integer
' Dim TM As String
'
' j = InStr(1, cmb_reqno, "$")
' rqno = req_chglen(Mid(Trim(cmb_reqno.Text), 1, j - 1))
'
' sel_rtimelog_qr(0) = rqno
' Set sel_rtimelog_rs = sel_rtimelog_qr.OpenResultset(1, 3)
'
' If Not sel_rtimelog_rs.EOF Then
'     If IsNull(sel_rtimelog_rs("SAMCOLDATE")) = False And Trim(sel_rtimelog_rs("SAMCOLDATE")) <> "" Then
'         msk_coll_dt.Text = sel_rtimelog_rs("SAMCOLDATE")
'         msk_coll_dt.Enabled = False
'     Else
'         msk_coll_dt.Enabled = True
'     End If
'     If IsNull(sel_rtimelog_rs("SAMCOLTIME")) = False And Trim(sel_rtimelog_rs("SAMCOLTIME")) <> "" Then
'
'        If InStr(1, sel_rtimelog_rs("SAMCOLTIME"), ":") = 2 Then
'            TM = "0" & Mid(sel_rtimelog_rs("SAMCOLTIME"), 1, 4)
'            msk_coll_tm.Text = TM
'            msk_coll_tm.Enabled = False
'        ElseIf InStr(1, sel_rtimelog_rs("SAMCOLTIME"), ":") = 3 Then
'            msk_coll_tm.Text = Mid(sel_rtimelog_rs("SAMCOLTIME"), 1, 5)
'            msk_coll_tm.Enabled = False
'        End If
'     Else
'         msk_coll_tm.Enabled = True
'     End If
' End If
''***** 22/03/11 ****************

'********* Shraddha Pandit 11/06/09 **************
If prev_flag = True Then
     cmd_view_Click
     getupperpanel_prev 'ADDED on 16/06/2009
End If
'*************************************************

'On Error GoTo MSG

clearlowerpanel2
'get what req is selected and depending on that change lower frame.
'FIRST GET MODALITY
' get detail of patient
If cmb_reqno.ListIndex > 0 And (cmb_reqno.Text) = "" Or Len(cmb_reqno.Text) < 1 Then
    MsgBox "Select Requisition First", vbOKOnly
    btnenabl "INI"
    Exit Sub
End If

'ADDED BY MILIND ON 16-10-06
Dim a As Integer
 a = InStr(2, cmb_reqno, "$")
mod_reqno = req_chglen(Mid(Trim(cmb_reqno), 1, a - 1)) 'put change value to this var so that it used later on
'mod_reqno = Split(Trim(cmb_reqno), "$", 2)
'**** Added by Shraddha on 17/02/10 Deselection for procodes *****
'**** Commented by Shraddha on 04/03/10 Again requirement came from microbiology to
'revert back.
'sel_dreqpc_qr(0) = Trim(mod_reqno)
'Set sel_dreqpc_rs = sel_dreqpc_qr.OpenResultset(1, 3)
'
'If sel_dreqpc_rs("PRO_CODE") = "G146" Then
'    MsgBox "Report can not be created for this requisitn", vbOKOnly
'    btnenabl "INI"
'    Exit Sub
'ElseIf sel_dreqpc_rs("PRO_CODE") = "G149" Then
'    MsgBox "Report can not be created for this requisitn", vbOKOnly
'    btnenabl "INI"
'    Exit Sub
'ElseIf sel_dreqpc_rs("PRO_CODE") = "G148" Then
'    MsgBox "Report can not be created for this requisitn", vbOKOnly
'    btnenabl "INI"
'    Exit Sub
'ElseIf sel_dreqpc_rs("PRO_CODE") = "G145" Then
'    MsgBox "Report can not be created for this requisitn", vbOKOnly
'    btnenabl "INI"
'    Exit Sub
'ElseIf sel_dreqpc_rs("PRO_CODE") = "G402" Then
'    MsgBox "Report can not be created for this requisitn", vbOKOnly
'    btnenabl "INI"
'    Exit Sub
'ElseIf sel_dreqpc_rs("PRO_CODE") = "G403" Then
'    MsgBox "Report can not be created for this requisitn", vbOKOnly
'    btnenabl "INI"
'    Exit Sub
'End If
'**************************


GETUSERLIST
'GETFINALLIST
If Opt_patient.Value Then
    If new_flag Then
        ''******ICU****************
            'VAR_REQNNO = req_chglen(cmb_reqno.Text)
        If SELICU_FLAG = False Then
            SELICU_QR(0) = mod_reqno ' VAR_REQNNO
            Set SELICU_RS = SELICU_QR.OpenResultset(1, 3)
            SELICU_FLAG = True
        Else
            SELICU_QR(0) = mod_reqno 'VAR_REQNNO
            SELICU_RS.MoreResults
            SELICU_RS.Requery
        End If
        
        If Not SELICU_RS.EOF Then
            If Trim(SELICU_RS("ICDS2")) = "YES" Then
                lblicu.Visible = True
            Else
                lblicu.Visible = False
            End If
        End If
            ''******ICU****************
        txt_addendum.Enabled = False
        cmb_addendumby.Enabled = False
       
    
       
        INSERT
        
        '********** Shraddha Pandit 23/10/08 ********
        RQN = Trim(cmb_reqno.Text)
        Dim b1 As Integer
        b1 = InStr(1, cmb_reqno, "$")
        RQN = req_chglen(Mid(Trim(cmb_reqno), 1, b1 - 1))

        Dim G_QR As rdoQuery
        Dim G_RS As rdoResultset
        Dim RS As rdoResultset
        str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO = ? "
        Set G_QR = rdb.CreateQuery("", str)
        G_QR(0) = RQN
        Set G_RS = G_QR.OpenResultset(1, 3)
        If G_RS.RowCount > 0 Then
            PRO_CODE = G_RS("PRO_CODE")
        
           If G_RS("PRO_CODE") = "G261" Or G_RS("PRO_CODE") = "G262" Or G_RS("PRO_CODE") = "G264" Or G_RS("PRO_CODE") = "G267" Then
           ' If G_RS("PRO_CODE") = "G261" Or G_RS("PRO_CODE") = "G265" Or G_RS("PRO_CODE") = "G262" Or G_RS("PRO_CODE") = "G264" Or G_RS("PRO_CODE") = "G267" Then
               ' frmSerology.Visible = True
                    frm_others.Visible = True
                str = "select TESTPERF from dislib.gzztest1 where sercode='" & PRO_CODE & "'"
                Set RS = rdb.OpenResultset(str, 1, 1)
                txtSerTestPerform.Text = RS("TESTPERF")
                RS.Close
                
                str = "select DISTINCT RESULT from dislib.gzztest1 where sercode='" & PRO_CODE & "'"
                Set RS = rdb.OpenResultset(str, 1, 1)
                While Not RS.EOF
                    cmbSerResult.AddItem (RS("RESULT"))
                    RS.MoveNext
                Wend
                RS.Close
                
                Select Case PRO_CODE
                    Case "G261"
                        str = "select DISTINCT PROCALC from dislib.gzztest1 where sercode='" & PRO_CODE & "'"
                        Set RS = rdb.OpenResultset(str, 1, 1)
                        While Not RS.EOF
                            cmbProcalci.AddItem Trim(RS("PROCALC"))
                            RS.MoveNext
                        Wend
                        RS.Close
                        
                        str = "select DISTINCT PROCAL_CN from dislib.gzztest1 where sercode='" & PRO_CODE & "' AND PROCAL_CN IS NOT NULL"
                        Set RS = rdb.OpenResultset(str, 1, 1)
                        While Not RS.EOF
                        
                            cmbProCalCon.AddItem Trim(RS("PROCAL_CN"))
                            RS.MoveNext
                        Wend
                        RS.Close
                        
                        str = "select DISTINCT TESTRANGE from dislib.gzztest1 where sercode='" & PRO_CODE & "' AND TESTRANGE IS NOT NULL"
                        Set RS = rdb.OpenResultset(str, 1, 1)
                        While Not RS.EOF
                            cmbSerRange.AddItem Trim(RS("TESTRANGE"))
                            RS.MoveNext
                        Wend
                        RS.Close
                    Case "G262"
                        
                    Case "G264"
                    
                    Case "G267"
                    
                End Select
            ElseIf G_RS("PRO_CODE") = "G151" Then
                Frame2.Visible = True
                MSFlexGrid1.Rows = 9
                'text1.Visible = False
                
                '***** 03/03/10 *****
                txt_test.SetFocus
                
                '********************
                
                MSFlexGrid1.clear
    
                'str = "SELECT * FROM DISLIB.MICRO_DRUG "
                str = "SELECT * FROM DISLIB.MICR_DRUG1 ORDER BY SRNO"
                Set SELDRUG_QR = rdb.CreateQuery("", str)
                Set SELDRUG_RS = SELDRUG_QR.OpenResultset(1, 3)
    
                MSFlexGrid1.ColWidth(0) = "1420"
                MSFlexGrid1.ColWidth(1) = "1420"
                MSFlexGrid1.ColWidth(2) = "1420"
                MSFlexGrid1.ColWidth(3) = "2000"
    
                MSFlexGrid1.TextMatrix(0, 0) = "DRUG"
                MSFlexGrid1.TextMatrix(0, 1) = "MIC"
                MSFlexGrid1.TextMatrix(0, 2) = "UOM"
                MSFlexGrid1.TextMatrix(0, 3) = "RESULT"
                
                
    
                Dim I As Integer
                countr = SELDRUG_RS.RowCount + 1
                MSFlexGrid1.Rows = countr
                I = 1

                While Not SELDRUG_RS.EOF
                    MSFlexGrid1.TextMatrix(I, 0) = Trim(SELDRUG_RS("drug_name"))
                    MSFlexGrid1.TextMatrix(I, 2) = "ug/ml"
                    SELDRUG_RS.MoveNext
                    I = I + 1
                Wend
                
'                '***** Shraddha 26/02/10 *****
'                Dim w As Integer
'                For w = 1 To 9
'                    MSFlexGrid1.TextMatrix(w, 2) = "ug/ml"
'                Next
'                '*****************************
            Else
                Frame2.Visible = False
            End If
        End If
    ''**************************************************
    '----------- Shraddha Pandit 12/11/08 -------------
    
        RQN = Trim(cmb_reqno.Text)
        Dim bn As Integer
        bn = InStr(1, cmb_reqno, "$")
        RQN = req_chglen(Mid(Trim(cmb_reqno), 1, bn - 1))

        Dim GB_QR As rdoQuery
        Dim GB_RS As rdoResultset
        str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO = ? "
        Set GB_QR = rdb.CreateQuery("", str)
        GB_QR(0) = RQN
        Set GB_RS = GB_QR.OpenResultset(1, 3)
        If GB_RS.RowCount > 0 Then
        'If GB_RS("PRO_CODE") = "G204" Or GB_RS("PRO_CODE") = "G206" Or GB_RS("PRO_CODE") = "G209" Or GB_RS("PRO_CODE") = "G111" Or GB_RS("PRO_CODE") = "G256" Or GB_RS("PRO_CODE") = "G268" Or GB_RS("PRO_CODE") = "G265" Or GB_RS("PRO_CODE") = "G266" Then
            If GB_RS("PRO_CODE") = "G204" Or GB_RS("PRO_CODE") = "G206" Or GB_RS("PRO_CODE") = "G209" Or GB_RS("PRO_CODE") = "G256" Or GB_RS("PRO_CODE") = "G268" Or GB_RS("PRO_CODE") = "G265" Or GB_RS("PRO_CODE") = "G266" Then
                Frame3.Visible = True
                q = True
                txt_btest.Text = ""
                txt_bresult.Text = ""
                txt_btest.SetFocus 'Added by Shraddha on 07/09/09
                '*****
'                frm_bacteria.Visible = False
'                frm_others.Visible = False
'                frm_clinic.Visible = False
'                frm_myco.Visible = False
                '******
            ElseIf GB_RS("PRO_CODE") = "G111" Then
            frm_others.Visible = True
            
            Else
                Frame3.Visible = False
                q = False
            End If
        End If
        
    '--------------------------------------------------
          
    ElseIf VIEW_FLAG Then
    
    '********************************************
    '-------- Shraddha Pandit 12/11/08 -------
        
        RQN = cmb_reqno.Text
        Dim bv As Integer
        bv = InStr(1, cmb_reqno, "$")
        RQN = req_chglen(Mid(Trim(cmb_reqno), 1, bv - 1))

        Dim GB_QR1 As rdoQuery
        Dim GB_RS1 As rdoResultset
        str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO = ? "
        Set GB_QR1 = rdb.CreateQuery("", str)
        GB_QR1(0) = RQN
        Set GB_RS1 = GB_QR1.OpenResultset(1, 3)
        If GB_RS1.RowCount > 0 Then
            If GB_RS1("PRO_CODE") = "G204" Or GB_RS1("PRO_CODE") = "G206" Or GB_RS1("PRO_CODE") = "G209" Or GB_RS1("PRO_CODE") = "G111" Or GB_RS1("PRO_CODE") = "G268" Or GB_RS1("PRO_CODE") = "G265" Or GB_RS1("PRO_CODE") = "G266" Or GB_RS1("PRO_CODE") = "G256" Then
                Frame3.Visible = True
                q = True
                str = "SELECT * FROM DISLIB.dmoprismer WHERE reqno = ? "
                Set ViewGBQR = rdb.CreateQuery("", str)

                gk = InStr(1, cmb_reqno, "$")

                ViewGBQR(0) = req_chglen(Mid(Trim(cmb_reqno), 1, gk - 1))
                Set ViewGBRS = ViewGBQR.OpenResultset(1, 3)
                If ViewGBRS.RowCount > 0 Then
                txt_btest.Text = Trim(ViewGBRS("OFINDING"))
                End If
                If ViewGBRS.RowCount > 0 Then
                txt_bresult.Text = Trim(ViewGBRS("REPORT"))
                End If
            ElseIf GB_RS1("PRO_CODE") = "G261" Or GB_RS1("PRO_CODE") = "G262" Or GB_RS1("PRO_CODE") = "G264" Or GB_RS1("PRO_CODE") = "G267" Then
'                str = "select * from dislib.gzzserolog where reqno='" & req_chglen(Mid(Trim(cmb_reqno), 1, bv - 1)) & "'"
'                Set ViewGBRS = rdb.OpenResultset(str, 1, 3)
'
'                frmSerology.Visible = True
'                txtSerTestPerform.Text = ViewGBRS("TESTPERF")
'                cmbSerResult.Text = ViewGBRS("Result")
'                cmbProcalci.Text = ViewGBRS("PROCALC")
'                cmbProCalCon.Text = ViewGBRS("PROCAL_CN")
'                cmbSerRange.Text = ViewGBRS("TESTRANGE")
'                txtInterpretation.Text = ViewGBRS("INTERPRETN")
'                txtMethod.Text = ViewGBRS("METHOD")
'                txtLimitation.Text = ViewGBRS("LIMITATION")
'                txtInBold.Text = ViewGBRS("FLDBLD")
            Else
                Frame3.Visible = False
                q = False
                txt_btest.Text = ""
                txt_bresult.Text = ""
            End If
            
       ' getresult_bacteriology
            
        End If
                    
        '-----------------------------------------
    '********************************************
    
    
        ''******ICU****************
        VAR_REQNNO = req_chglen(cmb_reqno.Text)
        If SELICU_FLAG = False Then
            SELICU_QR(0) = mod_reqno   'VAR_REQNNO
            Set SELICU_RS = SELICU_QR.OpenResultset(1, 3)
            SELICU_FLAG = True
        Else
            SELICU_QR(0) = VAR_REQNNO
            SELICU_RS.MoreResults
            SELICU_RS.Requery
        End If
        If Not SELICU_RS.EOF Then
            If Trim(SELICU_RS("ICDS2")) = "YES" Then
                lblicu.Visible = True
            Else
                lblicu.Visible = False
            End If
        End If
        ''******ICU****************
        RETRIVE
        
        '------------------------------------
        ''********** Shraddha Pandit 23/10/08 ********
        RQN = cmb_reqno.Text
        Dim b2 As Integer
        b2 = InStr(1, cmb_reqno, "$")
        RQN = req_chglen(Mid(Trim(cmb_reqno), 1, b2 - 1))

        Dim G_QR1 As rdoQuery
        Dim G_RS1 As rdoResultset
        str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO = ? "
        Set G_QR1 = rdb.CreateQuery("", str)
        G_QR1(0) = RQN
        Set G_RS1 = G_QR1.OpenResultset(1, 3)
        If G_RS1.RowCount > 0 Then
            If G_RS1("PRO_CODE") = "G151" Then
                Frame2.Visible = True
                MSFlexGrid1.Rows = 9
                MSFlexGrid1.ColWidth(0) = "1420"
                MSFlexGrid1.ColWidth(1) = "1420"
                MSFlexGrid1.ColWidth(2) = "1420"
                MSFlexGrid1.ColWidth(3) = "2000"

                MSFlexGrid1.TextMatrix(0, 0) = "DRUG"
                MSFlexGrid1.TextMatrix(0, 1) = "MIC"
                MSFlexGrid1.TextMatrix(0, 2) = "UOF"
                MSFlexGrid1.TextMatrix(0, 3) = "RESULT"
                
                '***** Shraddha 26/02/10 *****
                For I = 1 To 8
                    MSFlexGrid1.TextMatrix(I, 2) = "uM/L"
                Next
                '*****************************
                
                '***** Added by Shraddha on 04/07/2011 *****
                
                 str = "select * from DISLIB.dmycobacte where reqno=?"
                
                Set sel_comment_qr = rdb.CreateQuery("", str)
                
                sel_comment_qr(0) = mod_reqno
                Set sel_comment_rs = sel_comment_qr.OpenResultset(1, 3)
                'While Not sel_forupd6_rs.EOF
                If Not sel_comment_rs.EOF Then
                    If Trim(sel_comment_rs("comment")) <> "" Then
                        txt_MCcomment.Text = Trim(sel_comment_rs("comment"))
                    End If
                End If
                    'sel_forupd6_rs.MoveNext
                'Wend
                sel_comment_rs.Close
                '*******************************************
            Else
                Frame2.Visible = False
                txt_test.Text = ""
                txt_testresult.Text = ""
                txt_MCcomment.Text = ""
                MSFlexGrid1.clear
            End If
            '////////////////////////////
    
            str = "SELECT * FROM DISLIB.MICRO_TEST WHERE reqno = ? AND DRUG IN (SELECT DRUG_NAME FROM DISLIB.MICR_DRUG1 ORDER BY SRNO)"
            Set ViewMTQR = rdb.CreateQuery("", str)

            Dim K As Integer
            K = InStr(1, cmb_reqno, "$")

            ViewMTQR(0) = req_chglen(Mid(Trim(cmb_reqno), 1, K - 1))
            Set ViewMTRS = ViewMTQR.OpenResultset(1, 3)
            countr = ViewMTRS.RowCount
            Dim Iv As Integer
            Iv = 1
            If Not ViewMTRS.EOF Then
                txt_test.Text = Trim(ViewMTRS("Test"))
                txt_testresult.Text = Trim(ViewMTRS("Test_Res"))
                MSFlexGrid1.Rows = ViewMTRS.RowCount + 1
                While Iv < ViewMTRS.RowCount + 1
                    MSFlexGrid1.TextMatrix(Iv, 0) = Trim(ViewMTRS("Drug"))
                    MSFlexGrid1.TextMatrix(Iv, 1) = Trim(ViewMTRS("Mic"))
                    MSFlexGrid1.TextMatrix(Iv, 2) = Trim(ViewMTRS("Uom"))
                    MSFlexGrid1.TextMatrix(Iv, 3) = Trim(ViewMTRS("Result"))
                    Iv = Iv + 1
                    ViewMTRS.MoveNext
                Wend
            End If
        End If
        '------------------------------------
    
    End If
    
ElseIf opt_staff.Value Then 'FOR STAFF MEMBER
    'STAFF
    If new_flag Then
        '********** Shraddha Pandit 23/10/08 ********
        RQN = cmb_reqno.Text
        Dim N1 As Integer
        N1 = InStr(1, cmb_reqno, "$")
        RQN = req_chglen(Mid(Trim(cmb_reqno), 1, N1 - 1))

        Dim GS_QR As rdoQuery
        Dim GS_RS As rdoResultset
        str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO = ? "
        Set GS_QR = rdb.CreateQuery("", str)
        GS_QR(0) = RQN
        Set GS_RS = GS_QR.OpenResultset(1, 3)
        If GS_RS.RowCount > 0 Then
         PRO_CODE = GS_RS("PRO_CODE") '''added by charulata
            If GS_RS("PRO_CODE") = "G151" Then
                Frame2.Visible = True
                'text1.Visible = False
                MSFlexGrid1.clear
    
                str = "SELECT * FROM DISLIB.MICRO_DRUG"
                Set SELDRUG_QR = rdb.CreateQuery("", str)
                Set SELDRUG_RS = SELDRUG_QR.OpenResultset(1, 3)
    
                MSFlexGrid1.ColWidth(0) = "1420"
                MSFlexGrid1.ColWidth(1) = "1420"
                MSFlexGrid1.ColWidth(2) = "1420"
                MSFlexGrid1.ColWidth(3) = "2000"
    
                MSFlexGrid1.TextMatrix(0, 0) = "DRUG"
                MSFlexGrid1.TextMatrix(0, 1) = "MIC"
                MSFlexGrid1.TextMatrix(0, 2) = "UOM"
                MSFlexGrid1.TextMatrix(0, 3) = "RESULT"
    
                Dim S As Integer
                countr = SELDRUG_RS.RowCount + 1
                MSFlexGrid1.Rows = countr
                S = 1

                While Not SELDRUG_RS.EOF
                    MSFlexGrid1.TextMatrix(S, 0) = Trim(SELDRUG_RS("drug_name"))
                    SELDRUG_RS.MoveNext
                    S = S + 1
                Wend
            Else
                Frame2.Visible = False
            End If
        End If
        '***********************************
    
        '----- Shraddha Pandit 15/11/08 -----------
    
        RQN = Trim(cmb_reqno.Text)
        Dim bsn As Integer
        bsn = InStr(1, cmb_reqno, "$")
        RQN = req_chglen(Mid(Trim(cmb_reqno), 1, bsn - 1))

        Dim GBS_QR As rdoQuery
        Dim GBS_RS As rdoResultset
        str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO = ? "
        Set GBS_QR = rdb.CreateQuery("", str)
        GBS_QR(0) = RQN
        Set GBS_RS = GBS_QR.OpenResultset(1, 3)
        '**** vidya
        If Not GBS_RS.EOF Then
            If GBS_RS("PRO_CODE") = "G204" Or GBS_RS("PRO_CODE") = "G206" Or GBS_RS("PRO_CODE") = "G209" Or GBS_RS("PRO_CODE") = "G111" Or GBS_RS("PRO_CODE") = "G256" Or GBS_RS("PRO_CODE") = "G268" Or GBS_RS("PRO_CODE") = "G265" Or GBS_RS("PRO_CODE") = "G266" Then
                Frame3.Visible = True
                q = True
                txt_btest.Text = ""
                txt_bresult.Text = ""
            Else
                Frame3.Visible = False
                q = False
            End If
        End If
    End If
        
    '------------------------------------------
    '*********** Shraddha Pandit 12/11/08 ****************

    If VIEW_FLAG Then
    
        '-------- Shraddha Pandit 15/11/08 -----------
        RQN = cmb_reqno.Text
        Dim bsv As Integer
        bsv = InStr(1, cmb_reqno, "$")
        RQN = req_chglen(Mid(Trim(cmb_reqno), 1, bsv - 1))

        Dim GBS_QR1 As rdoQuery
        Dim GBS_RS1 As rdoResultset
        str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO = ? "
        Set GBS_QR1 = rdb.CreateQuery("", str)
        GBS_QR1(0) = RQN
        Set GBS_RS1 = GBS_QR1.OpenResultset(1, 3)
        If GBS_RS1.RowCount > 0 Then
            If GBS_RS1("PRO_CODE") = "G204" Or GBS_RS1("PRO_CODE") = "G206" Or GBS_RS1("PRO_CODE") = "G209" Or GBS_RS1("PRO_CODE") = "G111" Or GBS_RS1("PRO_CODE") = "G256" Or GBS_RS1("PRO_CODE") = "G268" Or GBS_RS1("PRO_CODE") = "G265" Or GBS_RS1("PRO_CODE") = "G266" Then
                Frame3.Visible = True
                q = True
                str = "SELECT * FROM DISLIB.dmoprismer WHERE reqno = ? "
                Set ViewGBQR = rdb.CreateQuery("", str)

                Dim gk1 As Integer
                gk1 = InStr(1, cmb_reqno, "$")

                ViewGBQR(0) = req_chglen(Mid(Trim(cmb_reqno), 1, gk1 - 1))
                Set ViewGBRS = ViewGBQR.OpenResultset(1, 3)
                
                txt_btest.Text = Trim(ViewGBRS("OFINDING"))
                txt_bresult.Text = Trim(ViewGBRS("REPORT"))
            
            Else
                Frame3.Visible = False
                q = False
                txt_btest.Text = ""
                txt_bresult.Text = ""
            
            End If
        End If
        '---------------------------------------------
        
        ''********** Shraddha Pandit 23/10/08 ********
        RQN = cmb_reqno.Text
        Dim a1 As Integer
        a1 = InStr(1, cmb_reqno, "$")
        RQN = req_chglen(Mid(Trim(cmb_reqno), 1, a1 - 1))

        Dim GV_QR As rdoQuery
        Dim GV_RS As rdoResultset
        str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO = ? "
        Set GV_QR = rdb.CreateQuery("", str)
        GV_QR(0) = RQN
        Set GV_RS = GV_QR.OpenResultset(1, 3)
        If GV_RS.RowCount > 0 Then
        
            If GV_RS("PRO_CODE") = "G151" Then
                Frame2.Visible = True
    
                MSFlexGrid1.ColWidth(0) = "1420"
                MSFlexGrid1.ColWidth(1) = "1420"
                MSFlexGrid1.ColWidth(2) = "1420"
                MSFlexGrid1.ColWidth(3) = "2000"

                MSFlexGrid1.TextMatrix(0, 0) = "DRUG"
                MSFlexGrid1.TextMatrix(0, 1) = "MIC"
                MSFlexGrid1.TextMatrix(0, 2) = "UOF"
                MSFlexGrid1.TextMatrix(0, 3) = "RESULT"
                
                '***** Shraddha 26/02/10 *****
                For I = 1 To 8
                    MSFlexGrid1.TextMatrix(I, 2) = "uM/L"
                Next
                '*****************************
                
            Else
                Frame2.Visible = False
                txt_test.Text = ""
                txt_testresult.Text = ""
                MSFlexGrid1.clear
                txt_MCcomment.Text = ""
            End If
        End If
        '////////////////////////////
    
        str = "SELECT * FROM DISLIB.MICRO_TEST WHERE reqno = ? "
        Set ViewMTQR = rdb.CreateQuery("", str)

        Dim v As Integer
        v = InStr(1, cmb_reqno, "$")

        ViewMTQR(0) = req_chglen(Mid(Trim(cmb_reqno), 1, v - 1))
        Set ViewMTRS = ViewMTQR.OpenResultset(1, 3)
        countr = ViewMTRS.RowCount
        Dim Pv As Integer
        Pv = 1
        If Not ViewMTRS.EOF Then
            txt_test.Text = Trim(ViewMTRS("Test"))
            txt_testresult.Text = Trim(ViewMTRS("Test_Res"))
            MSFlexGrid1.Rows = ViewMTRS.RowCount + 1
            While Pv < ViewMTRS.RowCount + 1
                MSFlexGrid1.TextMatrix(Pv, 0) = Trim(ViewMTRS("Drug"))
                MSFlexGrid1.TextMatrix(Pv, 1) = Trim(ViewMTRS("Mic"))
                MSFlexGrid1.TextMatrix(Pv, 2) = Trim(ViewMTRS("Uom"))
                MSFlexGrid1.TextMatrix(Pv, 3) = Trim(ViewMTRS("Result"))
                Pv = Pv + 1
                ViewMTRS.MoveNext
            Wend
        End If
    End If
    '-------
    STAFF ' Shraddha Pandit 15/11/08
    ''******************** End ***********************
    
End If 'END OF TOP IF
no_req = True


If frm_others.Visible = True And Mid(Trim(cmb_reqno.Text), 5, 2) = "MO" Then
'Add to combo

str = "select * from DISLIB.MICROMO"
Set sel_MOtemplet_qr = rdb.CreateQuery("", str)

Set sel_MOtemplet_rs = sel_MOtemplet_qr.OpenResultset(1, 3)

While Not sel_MOtemplet_rs.EOF
    cmb_motemplet.AddItem (Trim(sel_MOtemplet_rs("TESTNAME")))
    sel_MOtemplet_rs.MoveNext
Wend



End If
str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO = ? "
        Set GBS_QR = rdb.CreateQuery("", str)
        GBS_QR(0) = RQN
        Set GBS_RS = GBS_QR.OpenResultset(1, 3)
        '**** vidya
        If Not GBS_RS.EOF Then
            If GBS_RS("PRO_CODE") = "G204" Or GBS_RS("PRO_CODE") = "G262" Or GBS_RS("PRO_CODE") = "G206" Or GBS_RS("PRO_CODE") = "G209" Or GBS_RS("PRO_CODE") = "G111" Or GBS_RS("PRO_CODE") = "G256" Or GBS_RS("PRO_CODE") = "G268" Or GBS_RS("PRO_CODE") = "G265" Or GBS_RS("PRO_CODE") = "G266" Then
                Frame3.Visible = True
                q = True
                'txt_btest.Text = ""
               ' txt_bresult.Text = ""
            Else
                Frame3.Visible = False
                q = False
            End If
        End If
        
        str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO = ? "
        Set GBS_QR = rdb.CreateQuery("", str)
        GBS_QR(0) = RQN
        Set GBS_RS = GBS_QR.OpenResultset(1, 3)
        '**** vidya
        If Not GBS_RS.EOF Then
            If GBS_RS("PRO_CODE") = "G204" Or GBS_RS("PRO_CODE") = "G262" Or GBS_RS("PRO_CODE") = "G206" Or GBS_RS("PRO_CODE") = "G209" Or GBS_RS("PRO_CODE") = "G111" Or GBS_RS("PRO_CODE") = "G256" Or GBS_RS("PRO_CODE") = "G268" Or GBS_RS("PRO_CODE") = "G265" Or GBS_RS("PRO_CODE") = "G266" Then
                Frame3.Visible = True
                q = True
              '  txt_btest.Text = ""
              '  txt_bresult.Text = ""
            Else
                Frame3.Visible = False
                q = False
            End If
        End If
        
Exit Sub
MSG:
    MsgBox ERROR(Err)
End Sub



Private Sub cmbProCalCon_Click()
Dim RS As rdoResultset
str = "select * from dislib.gzztest1 where sercode='" & PRO_CODE & "' and PROCAL_CN ='" & Trim(cmbProCalCon.Text) & "'"
Set RS = rdb.OpenResultset(str, 1, 1)
If Not RS.EOF Then
        txtInterpretation.Text = Trim(RS("INTERPRETN"))
        txtMethod.Text = Trim(RS("method"))
        txtLimitation.Text = Trim(RS("Limitation"))
        txtInBold.Text = Trim(RS("FLDBLD"))
End If
RS.Close
End Sub

Private Sub cmbSerResult_Click()
Dim RS As rdoResultset
str = "select * from dislib.gzztest1 where sercode='" & PRO_CODE & "' and result='" & Trim(cmbSerResult.Text) & "'"
Set RS = rdb.OpenResultset(str, 1, 1)
If Not RS.EOF Then
        txtInterpretation.Text = Trim(RS("INTERPRETN"))
        txtMethod.Text = Trim(RS("method"))
        txtLimitation.Text = Trim(RS("Limitation"))
        txtInBold.Text = Trim(RS("FLDBLD"))
End If
RS.Close

End Sub

Private Sub cmbSerResult_LostFocus()
Dim RS As rdoResultset
str = "select * from dislib.gzztest1 where sercode='" & PRO_CODE & "' and result='" & Trim(cmbSerResult.Text) & "'"
Set RS = rdb.OpenResultset(str, 1, 1)
Select Case PRO_CODE
    
    Case "G261"
                        
    Case "G262"
                    
    Case "G264"
                    
    Case "G267"
        txtInterpretation.Text = Trim(RS("INTERPRETN"))
        txtMethod.Text = Trim(RS("method"))
        txtLimitation.Text = Trim(RS("Limitation"))
        txtInBold.Text = Trim(RS("FLDBLD"))
End Select
RS.Close
End Sub

Private Sub cmd_back_Click()
   '----mrunal
    str = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USR_ID='A" & usr_name & "' "
    Set REPWTRN_QR = rdb.CreateQuery("", str)
   Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)
    If REPWTRNRS.RowCount > 0 Then
    cmb_final.Text = REPWTRNRS("USERNAME")
    End If
    '----
 Label27.Caption = ""
CMD_NEW.Enabled = False
cmd_view.Enabled = False
If Report_microbiology1.Text2.Text = "" Then
Report_microbiology1.Text2.Text = "2"
End If
Dim micro_no
On Error Resume Next
If Report_microbiology1.Text2.Text <> "" Then
cmb_reqno.clear
If Report_microbiology1.Text2.Text <> "2" Then
a = micro_no
End If
a = a + 1
Dim micro
'micro = Split(CBCCOMMIT.cmb_srno.List(a), "&")
Report_microbiology1.show
Report_microbiology1.cmd_view_Click
Report_microbiology1.txt_spm.Text = cbccommit.MSFlexGrid1.TextMatrix(a, 4)
Report_microbiology1.MaskEdBox1.Text = cbccommit.MSFlexGrid1.TextMatrix(a, 2)
Report_microbiology1.cmb_reqno.AddItem (cbccommit.MSFlexGrid1.TextMatrix(a, 7))

 Dim b, REQNO
b = Len(Trim(cbccommit.MSFlexGrid1.TextMatrix(a, 7)))
 REQNO = cbccommit.MSFlexGrid1.TextMatrix(a, 7)
If b = 13 Then
REQNO = Left(cbccommit.MSFlexGrid1.TextMatrix(a, 7), 3) & "/" & Mid(cbccommit.MSFlexGrid1.TextMatrix(a, 7), 4, 2) & "/" & Mid(cbccommit.MSFlexGrid1.TextMatrix(a, 7), 6, 2) & "/" & Right(cbccommit.MSFlexGrid1.TextMatrix(a, 7), 6)
 End If

If b = 16 Then
REQNO = Left(cbccommit.MSFlexGrid1.TextMatrix(a, 7), 3) & Mid(cbccommit.MSFlexGrid1.TextMatrix(a, 7), 5, 2) & Mid(cbccommit.MSFlexGrid1.TextMatrix(a, 7), 8, 2) & Right(cbccommit.MSFlexGrid1.TextMatrix(a, 7), 6)
End If

If b = 27 Then
REQNO = Left(cbccommit.MSFlexGrid1.TextMatrix(a, 7), 3) & Mid(cbccommit.MSFlexGrid1.TextMatrix(a, 7), 5, 2) & Mid(cbccommit.MSFlexGrid1.TextMatrix(a, 7), 8, 2) & Mid(cbccommit.MSFlexGrid1.TextMatrix(a, 7), 11, 6)
End If
req = REQNO & "$" & cbccommit.MSFlexGrid1.TextMatrix(a, 1)
Report_microbiology1.cmb_reqno.SelText = REQNO & "$" & cbccommit.MSFlexGrid1.TextMatrix(a, 1)
Report_microbiology1.cmb_modality.Text = cbccommit.cmb_modality.Text
Report_microbiology1.Text2.Text = "2"
Report_microbiology1.cmb_reqno_Click

Else

btnenabl ("INI")
clearlowerpanel1
clearlowerpanel2
clearlowerpanel3
clearlowerpanel4
clearupperpanel
set_visible "All"
Frame2.Visible = False 'Shraddha Pandit 29/10/08
Frame3.Visible = False 'Shraddha Pandit 11/11/08
prev_flag = False 'Added on 11/06/09 by Shraddha
q = False
txt_test.Text = ""
txt_MCcomment.Text = ""
txt_testresult.Text = ""
MSFlexGrid1.clear
DIS_MAIN.show
Me.Hide
End If

End Sub

Public Sub cmd_cancel_Click()
   '----mrunal
    str = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USR_ID='A" & usr_name & "' "
    Set REPWTRN_QR = rdb.CreateQuery("", str)
   Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)
    If REPWTRNRS.RowCount > 0 Then
    cmb_final.Text = REPWTRNRS("USERNAME")
    End If
    '----
 Label27.Caption = ""
CMD_NEW.Enabled = True
cmd_view.Enabled = True
cmd_back.Enabled = True
cmd_previous.Enabled = True
Report_microbiology1.Text2.Text = ""

prev_flag = False 'Added on 11/06/09 by Shraddha

btnenabl ("INI")
clearupperpanel
clearlowerpanel1
clearlowerpanel2
clearlowerpanel3
clearlowerpanel4
clearlowerpanel5
set_flag "ALL"
txt_1.Visible = False
txt_2.Visible = False
txt_3.Visible = False
Frame2.Visible = False
Frame3.Visible = False 'Shraddha 11/11/08
frmSerology.Visible = False
txt_c_1.Text = ""

q = False
MSFlexGrid1.clear
txt_test.Text = ""
txt_MCcomment.Text = ""
txt_testresult.Text = ""
'CMD_NEW.SetFocus
set_visible "All"
lbl_bcomments.Left = 120
'lbl_bcomments.Top = 3720
lbl_bcomments.Top = 3720 'Shraddha 03/07/09
txt_comment.Left = 1080
'txt_comment.Top = 3720
txt_comment.Top = 4440 'Shraddha 03/07/09
cmb_dept.Enabled = True
cmb_modality.Enabled = True
cmb_dept.Enabled = True
cmb_reqno.Enabled = True
MaskEdBox1.Enabled = True
txt_c_modflg = False
set_flag "Ni"
cmb_antibiotic.Visible = False
MODE = ""

End Sub

Private Sub cmd_commit_Click()
   '----mrunal
    str = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USR_ID='A" & usr_name & "' "
    Set REPWTRN_QR = rdb.CreateQuery("", str)
   Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)
    If REPWTRNRS.RowCount > 0 Then
    cmb_final.Text = REPWTRNRS("USERNAME")
    End If
    '----
 Label27.Caption = ""
If Report_microbiology1.Text2.Text <> "" Then
   cmb_reqno.Text = req
 End If
'modified by mrunal--25_2
Dim j
 j = InStr(cmb_modality.Text, "$")
    txt_modcode.Text = Mid(cmb_modality.Text, j + 1, 2)
  
    If txt_modcode.Text = "BA" Then
    
If cmb_patho = "Yes" Then
If cmb_org1 = "" Then
If cmb_org2 = "" Then
If cmb_org3 = "" Then
 MsgBox "Atleast 1 organisms is needed", vbOKOnly
        cmb_org1.SetFocus
        Exit Sub
        End If
        End If
End If
End If
End If
'end 'Commented by Shraddha on 19/08/2011
''***** Shraddha 19/03/11 *****
''Check whether collection date and time has been entered or not.
'
'If new_flag = True Then
'
'    Dim a As Integer
'    Dim cur_date As Date
'    If (IsDate(msk_coll_dt.Text)) = False Then
'        MsgBox "Wrong Collection date"
'        Exit Sub
'    End If
'
'    If msk_coll_tm.Text = "__:__" Then
'        MsgBox "Please enter collection time"
'        Exit Sub
'    End If
'
'    cur_date = Format(DATE11, "yyyy-mm-dd")
'
'    If msk_coll_dt.Text > cur_date Then
'        MsgBox "Collection date can not be greater than todays date"
'        Exit Sub
'    End If
'
'    If Mid(Trim(msk_coll_tm.Text), 1, 2) > 24 Then
'        MsgBox ("Please enter time in 24 hours format..."), vbOKOnly
'        msk_coll_tm.Text = "__:__"
'        msk_coll_tm.SetFocus
'        Exit Sub
'    End If
'
'    If Mid(Trim(msk_coll_tm.Text), 4, 2) >= 60 Then
'        MsgBox ("Minutes can not exceeds 60..."), vbOKOnly
'        msk_coll_tm.Text = "__:__"
'        msk_coll_tm.SetFocus
'        Exit Sub
'    End If
'
'    a = InStr(1, cmb_reqno, "$")
'    mod_reqno = req_chglen(Mid(Trim(cmb_reqno), 1, a - 1))
'
'    upd_rtime_qr(0) = Format(msk_coll_dt.Text, "YYYY-MM-DD")
'    upd_rtime_qr(1) = Format(msk_coll_tm.Text, "HH:MM:SS")
'    'upd_rtime_qr(1) = Format(time1, "HH:MM:SS")
'    upd_rtime_qr(2) = mod_reqno
'    upd_rtime_qr.Execute
'
'End If
'
''***** 19/03/11 *****


prev_flag = False 'Added on 11/06/09
'***************** Shradha Pandit 04/06/09 *************************
If Len(Trim(cmb_enteryby)) = 0 Then
    MsgBox "The Enter by field is EMPTY.", vbCritical
    cmb_enteryby.SetFocus
    Exit Sub
End If

If Len(Trim(cmb_final)) = 0 Then
    MsgBox "The Finalized by field is EMPTY.", vbCritical
    cmb_final.SetFocus
    Exit Sub
End If
    
If Len(Trim(txt_c_report.Text)) >= 200 Then
    MsgBox "The Report Description Max. length is 200 Characters Only"
    txt_c_report.SetFocus
    Exit Sub
End If
'cmd_save_Click
'*******************************************************************


On Error GoTo MSG
Dim I As Integer
I = vbYes 'added on 04/06/09 by Shraddha and next 4 statements are commented
'i = MsgBox(" Do you want to Commit the Record?", vbYesNo, vbInformation)
'If i = vbNo Then
'    Exit Sub
'End If

If I = vbYes Then
On Error GoTo MSG:
 
''********SECURITY***************************
    CURUSRDET_QR(0) = Trim(cmb_final.Text)
    CURUSRDET_RS.MoreResults
    CURUSRDET_RS.Requery
    If CURUSRDET_RS.RowCount > 0 Then
    Dim STR11 As String
    STR11 = CURUSRDET_RS("USERNAME")
        
        If LOGINTYP > CURUSRDET_RS("usrtype") Then
          ''   no action as he can commit
        ElseIf LOGINTYP < CURUSRDET_RS("usrtype") Then
            MsgBox "Sorry! This Report needs  is to Commited by a Higher Level than  that of Yours"
            cmd_commit.SetFocus
            Exit Sub
        ElseIf LOGINTYP = CURUSRDET_RS("usrtype") Then
            'If Trim(LOGINUSRNAME) <> Trim(CURUSRDET_RS("username")) Then
                '  If loginusrid <> curusrdetrs("usr_id") Then
            '    MsgBox "Sorry! This Report is not to be Finalized by You"
             '   cmd_commit.SetFocus
              '  Exit Sub
            'End If
        End If
        If CURUSRDET_RS("COMMIT") <> "YES" Then
            MsgBox "Sorry! You are not authorised to Commit the Report of this Modality"
            cmd_commit.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Sorry! The Security detials of this user is not available check up in DISLNKUSR"
        cmd_commit.SetFocus
        Exit Sub
   End If
   
   
'*********************************************
   
''********SECURITY***************************
 
    rdb.BeginTrans
    upd_6_qr(0) = Left(Trim(cmb_final.Text), 30)
    upd_6_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    upd_6_qr(2) = "F"
    upd_6_qr(3) = "Y"
    
    upd_6_qr(4) = mod_reqno
    
   
    upd_6_qr.Execute
    rdb.CommitTrans
  'MODIFIED BY MRUNAL
 
    UPDRTIMEC_QR(0) = Format(txt_date.Text, "YYYY-MM-DD")
    UPDRTIMEC_QR(1) = Format(txt_time.Text, "HH:MM:SS")
                    
    UPDRTIMEC_QR(2) = mod_reqno
        UPDRTIMEC_QR.Execute
        ' END ---------
 
    UPDATERTIMELOG(0) = Format(txt_date.Text, "YYYY-MM-DD")
    'UPDATERTIMELOG(1) = Format(txt_time.Text, "HH:MM:SS")
    UPDATERTIMELOG(1) = Format(time1, "HH:MM:SS")
    UPDATERTIMELOG(2) = mod_reqno
    UPDATERTIMELOG.Execute
 
  Dim UPDATERTIMELOG1 As rdoQuery
  str = "UPDATE RISLIB.RTIMELOG SET COMREPDATE=current date,COMREPTIME=current time WHERE REQNNO=?"
  Set UPDATERTIMELOG1 = rdb.CreateQuery("", str)
      UPDATERTIMELOG1(0) = mod_reqno
      UPDATERTIMELOG1.Execute
 
 
    btnenabl "INI"
    If ins_mod_var = "ins" Then
        'MsgBox "Record has successfully been" & vbCrLf & "Committed."
         Label27.Caption = "Record Has Been Committed Successfully"
    End If
    If ins_mod_var = "mod" Then
      '  MsgBox "Record has successfully been" & vbCrLf & "Committed."
       Label27.Caption = "Record Has Been Committed Successfully"
    End If
End If
cmd_print.Enabled = True

'***************** Shradha Pandit 04/06/09 ***********************
'*******************************************************************
Label27.Caption = "Record Commited Sucessfully."
Exit Sub

MSG:
    rdb.RollbackTrans
    MsgBox Err.Description
    
End Sub

Private Sub cmd_decomit_Click()

Dim b As Integer
b = InStr(1, cmb_reqno, "$")
If Report_microbiology1.Text2.Text <> "" Then
If Report_microbiology1.Text2.Text = "2" Then
cmb_reqno.Text = req
Else
cmb_reqno.Text = Report_microbiology1.Text2.Text
 End If
  End If
mod_reqno = req_chglen(Mid(Trim(cmb_reqno), 1, b - 1)) 'put change value to this var so that it used later on

UPDRREQ_QR(0) = "T"
UPDRREQ_QR(1) = Null
UPDRREQ_QR(2) = mod_reqno
UPDRREQ_QR.Execute

cmb_reqno_Click
Label27.Caption = "Record Decommited Sucessfully."

End Sub

Private Sub CMD_DECOMMIT_Click()
   '----mrunal
    str = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USR_ID='A" & usr_name & "' "
    Set REPWTRN_QR = rdb.CreateQuery("", str)
   Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)
    If REPWTRNRS.RowCount > 0 Then
    cmb_final.Text = REPWTRNRS("USERNAME")
    End If
    '----
 Label27.Caption = ""
 
'Dim b As Integer
'b = InStr(1, cmb_reqno, "$")
'mod_reqno = req_chglen(Mid(Trim(cmb_reqno), 1, b - 1)) 'put change value to this var so that it used later on
'
'UPDRREQ_QR(0) = "T"
'UPDRREQ_QR(1) = Null
'If Report_microbiology.Text2.Text <> "" Then
'If Report_microbiology.Text2.Text = "2" Then
''cmb_reqno.Text = req
'UPDRREQ_QR(2) = mod_reqno
'Else
'UPDRREQ_QR(2) = Report_microbiology.Text2.Text
' End If
'  End If
  
'UPDRREQ_QR(2) = mod_reqno
'UPDRREQ_QR.Execute
'
'cmb_reqno_Click
cbccommit.show
Report_microbiology1.Hide
cbccommit.cmd_showrpt_Click
End Sub

Private Sub cmd_fullreport_Click()
End
End Sub

Private Sub cmd_modify_Click()
   '----mrunal
    str = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USR_ID='A" & usr_name & "' "
    Set REPWTRN_QR = rdb.CreateQuery("", str)
   Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)
    If REPWTRNRS.RowCount > 0 Then
    cmb_final.Text = REPWTRNRS("USERNAME")
    End If
    '----
 Label27.Caption = ""
cmd_back.Enabled = False
cmd_previous.Enabled = False
'modified by mrunal--25_2
If Report_microbiology1.Text2.Text <> "" Then
If Report_microbiology1.Text2.Text = "2" Then
cmb_reqno.Text = req
Else
   cmb_reqno.Text = Report_microbiology1.Text2.Text
 End If
  End If
 I = InStr(cmb_modality.Text, "$")
    txt_modcode.Text = Mid(cmb_modality.Text, I + 1, 2)
  
    If txt_modcode.Text = "BA" Then
 If cmb_org1.Visible = True Then
If cmb_patho = "Yes" Then
If cmb_org1 = "" Then
If cmb_org2 = "" Then
If cmb_org3 = "" Then
 MsgBox "Atleast 1 organisms is needed", vbOKOnly
 If cmb_org1.Enabled = True Then
  'cmb_org1.SetFocus
 Else
 
 End If
       cmb_org1.Enabled = True
        Exit Sub
        End If
        End If
End If
End If
End If
End If
'end
btnenabl ("NEW")
MODE = "MODIFY"
cmb_antibiotic.clear
GET_ANTIBIOTIC

prev_flag = False 'Added by Shraddha on 11/06/09

'grid_dmbanti.Enabled = True

If Len(Trim(MaskEdBox1.Text)) = 0 Then
    MsgBox "Case Number should not be empty!", vbCritical

    Exit Sub
End If

If Len(Trim(cmb_reqno.Text)) = 0 Then
    MsgBox "Requisition Number should not be empty!", vbCritical
    If cmb_reqno.Enabled = True Then
        cmb_reqno.SetFocus
        SendKeys "^{f4}"
    End If
    Exit Sub
End If

'If REFDRFLAG = False Then
    REFDR_QR(0) = MaskEdBox1.FormattedText
    REFDR_QR(1) = req_chglen(cmb_reqno.Text)
   
    Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
    'REFDRFLAG = True
'Else
'    REFDR_QR(0) = MaskEdBox1.FormattedText
'    REFDR_QR(1) = req_chglen(cmb_reqno.Text)
'    REFDR_RS.MoreResults
'    REFDR_RS.Requery
'End If
If REFDR_RS.RowCount > 0 Then
    If Trim(REFDR_RS("reqnstat")) = "T" And Trim(REFDR_RS("COMMITFLAG")) = "F" Then
    ADDENDUMFALG = True
    End If
End If
'If ADDENDUMFALG = True Then
'check for commit flag status
sel_ris_qr(0) = mod_reqno
Set sel_ris_rs = sel_ris_qr.OpenResultset(1, 3)
While Not sel_ris_rs.EOF
    If Trim(sel_ris_rs(0)) = "Y" Then
        MsgBox "Sorry Record is Already Commited", vbOKOnly
        Exit Sub
    Else
        set_flag "MOD"
    End If
    sel_ris_rs.MoveNext
Wend

enable_field

'***** Shraddha 12/05/11 *****
If mod_flag = True Then
    grid_c_anti.Enabled = True
'    cmb_antiboitic_mb.Move grid_c_anti.Left + grid_c_anti.CellLeft, grid_c_anti.Top + grid_c_anti.CellTop, grid_c_anti.CellWidth
'    grid_c_anti.col = 1
'    txt_c_1.Move grid_c_anti.Left + grid_c_anti.CellLeft, grid_c_anti.Top + grid_c_anti.CellTop, grid_c_anti.CellWidth
   GET_ANTIBIOTIC_MB
End If

'********* 12/05/2011 ********

End Sub

Public Sub cmd_new_Click()
   '----mrunal
    str = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USR_ID='A" & usr_name & "' "
    Set REPWTRN_QR = rdb.CreateQuery("", str)
   Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)
    If REPWTRNRS.RowCount > 0 Then
    cmb_final.Text = REPWTRNRS("USERNAME")
    End If
    '----
 Label27.Caption = ""
cmd_back.Enabled = False
cmd_previous.Enabled = False
cmb_final.Enabled = False
cmd_commit.Visible = False
cmd_save.Visible = True

'lbl_colldt.Visible = True
'lbl_colltm.Visible = True
'msk_coll_dt.Visible = True
'msk_coll_tm.Visible = True

grid_c_anti.clear
prev_flag = False 'Added by Shraddha on 11/06/09

txt_c_1.Text = ""

'grid_dmbanti.Enabled = True
MODE = "NEW"
ins_mod_var = "ins"
btnenabl ("NEW")
    cmb_enteryby.Text = cmb_enteryby.List(0)
    cmb_enteryby.Enabled = False
     'cmb_final.Text = cmb_final.List(0)
    Set sel_deptrs = sel_deptqr.OpenResultset(1, 3)
    While Not sel_deptrs.EOF
        cmb_dept.AddItem Trim(sel_deptrs("dept_name")) & "$" & Trim(sel_deptrs("dept_code"))
        sel_deptrs.MoveNext
    Wend
    For I = 0 To (cmb_dept.ListCount - 1)
        If Trim(cmb_dept.List(I)) = "MICROBIOLOGY$GZZ" Then
                cmb_dept.ListIndex = I
        End If
    Next I
     If Trim(cmb_dept) = "" Then
        Exit Sub
    End If
    'copied and pasted
    I = InStr(cmb_dept.Text, "$")
    txt_deptcode.Text = Mid(cmb_dept.Text, I + 1, 3)
    sel_moadlityqr(0) = Trim(txt_deptcode.Text)
    Set sel_modalityrs = sel_moadlityqr.OpenResultset(1, 3)
    If sel_modalityrs.RowCount > 0 Then
        cmb_modality.clear
        While Not sel_modalityrs.EOF
            cmb_modality.AddItem Trim(sel_modalityrs("modlitydes")) & "$" & Trim(sel_modalityrs("modalitycd"))
            sel_modalityrs.MoveNext
        Wend
        sel_modalityrs.Close
    Else
        sel_modalityrs.Close
    End If
    '************** Shraddha Pandit 04/06/09 ************
    cmb_modality.Text = cmb_modality.List(0)
    cmb_modality_Click
    '****************************************************
sel_deptrs.Close
set_flag "NEW"
Opt_patient.Value = True
cmd_modify.Enabled = False
cmd_commit.Enabled = False
Frame2.Visible = False
Frame3.Visible = False 'Shraddha 11/11/08
q = False
MSFlexGrid1.clear
txt_test.Text = ""
txt_MCcomment.Text = ""
txt_testresult.Text = ""
txt_btest.Text = ""
txt_bresult.Text = ""
grid_dmbanti.Rows = 2

'MaskEdBox1.SetFocus ''13-AUG-09 SHWETA
End Sub

Private Sub cmd_previous_Click()
   '----mrunal
    str = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USR_ID='A" & usr_name & "' "
    Set REPWTRN_QR = rdb.CreateQuery("", str)
   Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)
    If REPWTRNRS.RowCount > 0 Then
    cmb_final.Text = REPWTRNRS("USERNAME")
    End If
    '----
 Label27.Caption = ""
CMD_NEW.Enabled = False
cmd_view.Enabled = False
' prev_flag = True
'MousePointer = vbHourglass
'Dim sel_prevreqn_qr As rdoQuery
'Dim sel_prevreqn_rs As rdoResultset
'Dim sql As String
''COMREPTIME
''COMREPDATE
'
'sql = "select * from RISLIB.rtimelog where COMREPDATE='" & Format(date11, "YYYY-MM-DD") & "' and reqnno like 'GZZ%' order by COMREPTIME desc"
'Set sel_prevreqn_qr = rdb.CreateQuery("", sql)
'
'Set sel_prevreqn_rs = sel_prevreqn_qr.OpenResultset(1, 3)
'
'While Not sel_prevreqn_rs.EOF
'    GZZREQPATH_QR(0) = sel_prevreqn_rs("reqNno")
'    Set GZZREQPATH_RS = GZZREQPATH_QR.OpenResultset(1, 3)
'    cmb_reqno.AddItem req_chglen(sel_prevreqn_rs("reqNno")) + "$" + GZZREQPATH_RS("PATHNO")
'    sel_prevreqn_rs.MoveNext
'Wend
'MousePointer = vbNormal
MousePointer = vbNormal
If Report_microbiology1.Text2.Text = "" Then
Report_microbiology1.Text2.Text = "2"
End If
Dim micro_no
If Report_microbiology1.Text2.Text <> "" Then
cmb_reqno.clear
If Report_microbiology1.Text2.Text <> "2" Then
a = micro_no
End If
a = a - 1
Dim micro
'micro = Split(CBCCOMMIT.cmb_srno.List(a), "&")
Report_microbiology1.show
Report_microbiology1.cmd_view_Click
Report_microbiology1.txt_spm.Text = cbccommit.MSFlexGrid1.TextMatrix(a, 4)
Report_microbiology1.MaskEdBox1.Text = cbccommit.MSFlexGrid1.TextMatrix(a, 2)
Report_microbiology1.cmb_reqno.AddItem (cbccommit.MSFlexGrid1.TextMatrix(a, 7))

 Dim b, REQNO
b = Len(Trim(cbccommit.MSFlexGrid1.TextMatrix(a, 7)))
 REQNO = cbccommit.MSFlexGrid1.TextMatrix(a, 7)
If b = 13 Then
REQNO = Left(cbccommit.MSFlexGrid1.TextMatrix(a, 7), 3) & "/" & Mid(cbccommit.MSFlexGrid1.TextMatrix(a, 7), 4, 2) & "/" & Mid(cbccommit.MSFlexGrid1.TextMatrix(a, 7), 6, 2) & "/" & Right(cbccommit.MSFlexGrid1.TextMatrix(a, 7), 6)
 End If

If b = 16 Then
REQNO = Left(cbccommit.MSFlexGrid1.TextMatrix(a, 7), 3) & Mid(cbccommit.MSFlexGrid1.TextMatrix(a, 7), 5, 2) & Mid(cbccommit.MSFlexGrid1.TextMatrix(a, 7), 8, 2) & Right(cbccommit.MSFlexGrid1.TextMatrix(a, 7), 6)
End If

If b = 27 Then
REQNO = Left(cbccommit.MSFlexGrid1.TextMatrix(a, 7), 3) & Mid(cbccommit.MSFlexGrid1.TextMatrix(a, 7), 5, 2) & Mid(cbccommit.MSFlexGrid1.TextMatrix(a, 7), 8, 2) & Mid(cbccommit.MSFlexGrid1.TextMatrix(a, 7), 11, 6)
End If
req = REQNO & "$" & cbccommit.MSFlexGrid1.TextMatrix(a, 1)
Report_microbiology1.cmb_reqno.SelText = REQNO & "$" & cbccommit.MSFlexGrid1.TextMatrix(a, 1)
Report_microbiology1.cmb_modality.Text = cbccommit.cmb_modality.Text
Report_microbiology1.Text2.Text = "2"
Report_microbiology1.cmb_reqno_Click

Else
btnenabl ("INI")
clearlowerpanel1
clearlowerpanel2
clearlowerpanel3
clearlowerpanel4
clearupperpanel
set_visible "All"
Frame2.Visible = False 'Shraddha Pandit 29/10/08
Frame3.Visible = False 'Shraddha Pandit 11/11/08
prev_flag = False 'Added on 11/06/09 by Shraddha
q = False
txt_test.Text = ""
txt_MCcomment.Text = ""
txt_testresult.Text = ""
MSFlexGrid1.clear
DIS_MAIN.show
Me.Hide
End If


End Sub

Private Sub cmd_print_Click()
 If Report_microbiology1.Text2.Text <> "" Then
 If Report_microbiology1.Text2.Text = "2" Then
    cmb_reqno.Text = req
Else
     cmb_reqno.Text = Report_microbiology1.Text2.Text
End If
End If
If Len(Trim(cmb_reqno.Text)) > 0 Then
    Dim VAR1
    VAR1 = mod_reqno & "~GUEST"
    Dim strLink As String
    'strLink = "http://intranet.tmc.gov.in/emrweb/default.asp" '10.1.4.9
    strLink = "http://intranet.tmc.gov.in/EMRREPORTS/allreports.aspx?reqno=" & VAR1 & ""
    Shell "c:\program files\internet explorer\iexplore " & strLink
Else
    MsgBox "Please Enter The Case No."
End If
Exit Sub
'On Error GoTo MSG
Dim TEMP_RS As rdoResultset
prev_flag = False 'Added on 11/06/09

SEL_DMUL_QR(0) = mod_reqno
Set SEL_DMUL_RS = SEL_DMUL_QR.OpenResultset(1, 3)

Dim UI As String
UI = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
'----- Shraddha Pandit 14/11/08 -----
If frmSerology.Visible = True Then
        For I = 0 To 14
            CrystalReport2.Formulas(I) = ""
        Next I
        
        CrystalReport2.SelectionFormula = "{zgzzserlf.REQNO} = '" & mod_reqno & "'"
        CrystalReport2.ReportFileName = "c:\DISREPORTSEClf\GZZSER_REPORT.RPT"
        'CrystalReport2.ReportFileName = "c:\DISREPORTSEClf\other_report.rpt"
        CrystalReport2.Formulas(0) = "proFIN = 'MICROBIOLOGY REPORT'"
        CrystalReport2.Formulas(1) = "ENT = '" & "Entered By   " & Trim(cmb_enteryby.Text) & "'"
        CrystalReport2.Formulas(3) = "bedno='" & txt_bedno.Text & "'"
        CrystalReport2.Formulas(4) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
        CrystalReport2.Formulas(5) = "pathnumber='" & txt_microno.Text & "'"
        CrystalReport2.Formulas(6) = "NAM='" & txt_name.Text & "'"
        CrystalReport2.Formulas(7) = "STDOB='" & txt_sex.Text & " /" & txt_age.Text & "'"
        CrystalReport2.Formulas(8) = "ST='" & txt_category.Text & " /" & txt_status.Text & "'"
        CrystalReport2.Formulas(9) = "USRID = '" & UI & "'"
        Set TEMP_RS = rdb.OpenResultset("select * from dislib.gzzserolog WHERE  REQNO = '" & mod_reqno & "'", 1, 1)
        If Not TEMP_RS.EOF Then
            If Trim(TEMP_RS("Result")) = "" Then CrystalReport2.Formulas(10) = "flgresult='false'"
            If Trim(TEMP_RS("INTERPRETN")) = "" Then CrystalReport2.Formulas(11) = "flginter='false'"
            If Trim(TEMP_RS("METHOD")) = "" Then CrystalReport2.Formulas(12) = "flgMethod='false'"
            If Trim(TEMP_RS("LIMITATION")) = "" Then CrystalReport2.Formulas(13) = "flgLimitation='false'"
        End If
        If Not Trim(txt_ward.Text) = "" Then
            If InStr(1, txt_ward.Text, "WARD") = 0 Then
                CrystalReport2.Formulas(14) = "ward='" & txt_ward.Text & " /" & txt_bedno.Text & "'"
            Else
                CrystalReport2.Formulas(14) = "ward='" & Trim(Mid(txt_ward.Text, 1, InStr(1, txt_ward.Text, "WARD") - 1)) & "'"
            End If
        Else
            CrystalReport2.Formulas(14) = "ward=''"
        End If
        CrystalReport2.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport2.WindowState = crptMaximized
        CrystalReport2.Action = 1
ElseIf Frame3.Visible = True Or q = True Then
    If Opt_patient = True Then
        CrystalReport2.SelectionFormula = "{RREQUISITN.REQNNO} = '" & mod_reqno & "'"
        CrystalReport2.ReportFileName = "c:\DISREPORTSEC\bacteriology_con.rpt"
        CrystalReport2.Formulas(0) = "proFIN = 'MICROBIOLOGY REPORT'"
        CrystalReport2.Formulas(1) = "ENT = '" & "Entered By   " & Trim(cmb_enteryby.Text) & "'"
        CrystalReport2.Formulas(2) = "patho = '" & "Yes" & "'"
        CrystalReport2.Formulas(3) = "bedno='" & txt_bedno.Text & "'"
        CrystalReport2.Formulas(4) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
        CrystalReport2.Formulas(5) = "pathnumber='" & txt_microno.Text & "'"
        CrystalReport2.Formulas(6) = "NAM='" & txt_name.Text & "'"
        CrystalReport2.Formulas(7) = "STDOB='" & txt_sex.Text & " /" & txt_age.Text & "'"
        CrystalReport2.Formulas(8) = "ST='" & txt_category.Text & " /" & txt_status.Text & "'"
        CrystalReport2.Formulas(9) = "USRID = '" & UI & "'"
        Set TEMP_RS = rdb.OpenResultset("SELECT * FROM DISLIB.DMULPROCE WHERE  REQNO = '" & mod_reqno & "'")
        If TEMP_RS.EOF = False Then
            If Trim(UCase(TEMP_RS("PRO_CODE"))) = "G204" Or Trim(UCase(TEMP_RS("PRO_CODE"))) = "G205" Or Trim(UCase(TEMP_RS("PRO_CODE"))) = "G206" Then
            CRYSTALREPORT1.Formulas(10) = "TXTPRIMARY_STAIN = ''"
            Else
            CRYSTALREPORT1.Formulas(10) = "TXTPRIMARY_STAIN = 'Primary Smear (Gram Stain) '"
            End If
            If TEMP_RS("PRO_CODE") = "G256" Then
            CrystalReport2.Formulas(10) = "PComment='Yes'"
            Else
                CrystalReport2.Formulas(10) = "PComment=''"
            End If
        End If
                  
        CrystalReport2.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport2.WindowState = crptMaximized
        CrystalReport2.Action = 1
    End If
    
    If opt_staff = True Then
        CrystalReport2.SelectionFormula = "{RREQUISITN.REQNNO} = '" & mod_reqno & "'"
        CrystalReport2.ReportFileName = "c:\DISREPORTSEClf\bacteriology_con_staff.rpt"
        CrystalReport2.Formulas(0) = "proFIN = 'MICROBIOLOGY REPORT'"
        CrystalReport2.Formulas(1) = "ENT = '" & "Entered By   " & Trim(cmb_enteryby.Text) & "'"
        CrystalReport2.Formulas(2) = "patho = '" & "Yes" & "'"
        CrystalReport2.Formulas(3) = "bedno='" & txt_bedno.Text & "'"
        CrystalReport2.Formulas(4) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
        CrystalReport2.Formulas(5) = "pathnumber='" & txt_microno.Text & "'"
        CrystalReport2.Formulas(6) = "NAM='" & txt_name.Text & "'"
        CrystalReport2.Formulas(7) = "STDOB='" & txt_sex.Text & " /" & txt_age.Text & "'"
        CrystalReport2.Formulas(8) = "ST='Staff /" & txt_status.Text & "'"
        CrystalReport2.Formulas(9) = "USRID = '" & UI & "'"
        If SEL_DMUL_RS("PRO_CODE") = "G256" Then
            CrystalReport2.Formulas(10) = "PComment='Yes'"
        Else
            CrystalReport2.Formulas(10) = "PComment=''"
        End If
        CrystalReport2.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport2.WindowState = crptMaximized
        CrystalReport2.Action = 1
    End If
    '------------------------------------
    '******** Shraddha Pandit 29/10/08 *******************
ElseIf Frame2.Visible = True Then
    Dim S1 As Integer
    If Opt_patient = True Then
    For S1 = 0 To 20
        CrystalReport2.Formulas(S1) = ""
    Next S1
    
        CrystalReport2.ReportFileName = "C:\DISREPORTSEC\MICRO_MYCO.rpt"
        CrystalReport2.WindowTitle = "Microbiology Report"
        
        CrystalReport2.SelectionFormula = "{micro_test.REQNO} = '" & mod_reqno & "'"
        CrystalReport2.Formulas(0) = "name ='" & txt_name & "'"
        CrystalReport2.Formulas(1) = "sex ='" & txt_sex & "'"
        CrystalReport2.Formulas(2) = "age ='" & txt_age & "'"
        CrystalReport2.Formulas(3) = "USRID = '" & UI & "'"
        CrystalReport2.Formulas(4) = "comment = '" & Trim(txt_MCcomment.Text) & "'"
        CrystalReport2.Destination = crptToWindow 'SHWETA 13-AUG-09
        'CrystalReport2.Destination = crptToPrinter 'SHRADDHA 09-11-09
        'CrystalReport2.Destination = crptToPrinter
        CrystalReport2.WindowShowPrintBtn = True
        CrystalReport2.WindowShowRefreshBtn = True
        CrystalReport2.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport2.WindowState = crptMaximized
        CrystalReport2.Action = 1
    End If
    
    If opt_staff = True Then
        CrystalReport2.WindowTitle = "Microbiology Report"
        CrystalReport2.ReportFileName = "C:\DISREPORTSEC\MICRO_MYCO_STF.rpt"
        
        CrystalReport2.SelectionFormula = "{micro_test.REQNO} = '" & mod_reqno & "'"
        CrystalReport2.Formulas(0) = "ccname ='" & txt_name & "'"
        CrystalReport2.Formulas(1) = "sex ='" & txt_sex & "'"
        CrystalReport2.Formulas(2) = "age ='" & txt_age & "'"
        CrystalReport2.Formulas(3) = "USRID = '" & UI & "'"
        
        'CrystalReport2.Destination = crptToWindow
         CrystalReport2.Destination = crptToPrinter
        CrystalReport2.WindowShowPrintBtn = True
        CrystalReport2.WindowShowRefreshBtn = True
        CrystalReport2.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport2.WindowState = crptMaximized
        CrystalReport2.Action = 1
    End If
        'Shraddha Pandit 23/01/09
ElseIf frm_others.Visible = True And (Not SEL_DMUL_RS.EOF) And SEL_DMUL_RS("PRO_CODE") = "G252" Then
    
        CRYSTALREPORT1.SelectionFormula = "{REPORTOTLF.REQNO} = '" & mod_reqno & "'"
        'CrystalReport1.ReportFileName = "c:\DISREPORTSEC\Other_report.rpt"
        'Commented and added by Shraddha on 15/09/10
        CRYSTALREPORT1.ReportFileName = "c:\DISREPORTSECLF\Other_report.rpt"
        CRYSTALREPORT1.Formulas(1) = "profin = 'MICROBIOLOGY REPORT'"
        CRYSTALREPORT1.Formulas(2) = "ENT = '" & "Entered By   " & Trim(cmb_enteryby.Text) & "'"
        CRYSTALREPORT1.Formulas(3) = "USRID = '" & str & "'"
        CRYSTALREPORT1.Formulas(13) = "bedno='" & txt_bedno.Text & "'"
        CRYSTALREPORT1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
        CRYSTALREPORT1.Formulas(6) = "pathnumber='" & txt_microno.Text & "'" '17/11/08 7->6
        If Not Trim(txt_ward.Text) = "" Then
            If InStr(1, txt_ward.Text, "WARD") = 0 Then
                CRYSTALREPORT1.Formulas(7) = "ward='" & txt_ward.Text & " /" & txt_bedno.Text & "'"
            Else
                CRYSTALREPORT1.Formulas(7) = "ward='" & Trim(Mid(txt_ward.Text, 1, InStr(1, txt_ward.Text, "WARD") - 1)) & "'"
            End If
        Else
            CRYSTALREPORT1.Formulas(7) = "ward="
        End If
        
        CRYSTALREPORT1.Formulas(8) = "NAM='" & txt_name.Text & "'"
        CRYSTALREPORT1.Formulas(9) = "STDOB='" & txt_sex.Text & " /" & txt_age.Text & "'"
        If Opt_patient.Value Then
            CRYSTALREPORT1.Formulas(10) = "ST='" & txt_category.Text & " /" & txt_status.Text & "'"
        Else
            CRYSTALREPORT1.Formulas(10) = "ST='Staff /" & txt_status.Text & "'"
        End If
        
        
        'CRYSTALREPORT1.Destination = crptToWindow 'Commented by Shraddha on 05/06/09
        CRYSTALREPORT1.Destination = crptToPrinter
        CRYSTALREPORT1.WindowState = crptMaximized
        
        CRYSTALREPORT1.WindowShowPrintBtn = True
        CRYSTALREPORT1.WindowShowRefreshBtn = True
        CRYSTALREPORT1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CRYSTALREPORT1.Action = 1
  
Else
    '**************************************
Dim prepath As String
Dim I1 As Integer
DoEvents
'check
For I1 = 0 To 20 'Increased from 10 to 20 on 05/08/09 by Shraddha Pandit
    CRYSTALREPORT1.Formulas(I1) = ""
Next I1
If no_req Then 'check for entered selected req no.
    
        ''*******icu****************
'    If Opt_patient.Value = True Then
'        If SELICU_FLAG = False Then
'            SELICU_QR(0) = mod_reqno
'            Set SELICU_RS = SELICU_QR.OpenResultset(1, 3)
'            SELICU_FLAG = True
'        Else
'            SELICU_QR(0) = mod_reqno
'            SELICU_RS.MoreResults
'            SELICU_RS.Requery
'        End If
'        If Not SELICU_RS.EOF Then
'            If SELMEMO_FLAG = False Then
'                SELMEMO_QR(0) = mod_reqno
'                Set SELMEMO_RS = SELMEMO_QR.OpenResultset(1, 3)
'                SELMEMO_FLAG = True
'            Else
'                SELMEMO_QR(0) = mod_reqno
'                SELMEMO_RS.MoreResults
'                SELMEMO_RS.Requery
'            End If
'        End If
'        If Not SELMEMO_RS.EOF Then
'            If IsNull(SELMEMO_RS("MEMONO")) And Trim(SELMEMO_RS("PAT_CATG")) <> "NC" Then
'                MsgBox " Please Prepare the Memo !!!", vbInformation, "DIS"
'                Exit Sub
'            End If
'        End If
'    End If
        ''*******icu****************
    sel_micro_qr(0) = mod_reqno
    Set sel_micro_rs = sel_micro_qr.OpenResultset(1, 3)
    While Not sel_micro_rs.EOF
        If Trim(sel_micro_rs("pathno")) <> Trim(txt_microno.Text) Then
             prepath = prepath & sel_micro_rs("pathno")
        End If
        sel_micro_rs.MoveNext
    Wend
    sel_ris_qr(0) = mod_reqno
    Set sel_ris_rs = sel_ris_qr.OpenResultset(1, 3)
    str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
      '  If ADDENDUMFALG = False Then
            If Trim(txt_modcode.Text) = "BA" And Frame3.Visible = False Then 'for case no. 1(bacteriology)
                If cmb_patho.Text = "Yes" Then
                    CRYSTALREPORT1.SelectionFormula = "{bacterolf.REQNNO} = '" & mod_reqno & "'"
                    'SHWETA 06-SEPT-2010
                    CRYSTALREPORT1.ReportFileName = "c:\DISREPORTSECLF\bacteriology_report.rpt"
                    CRYSTALREPORT1.Formulas(1) = "proFIN = 'MICROBIOLOGY REPORT'"
                    CRYSTALREPORT1.Formulas(2) = "ENT = '" & "Entered By   " & Trim(cmb_enteryby.Text) & "'"
                    CRYSTALREPORT1.Formulas(3) = "USRID = '" & str & "'"
                    CRYSTALREPORT1.Formulas(4) = "patho = '" & "Yes" & "'"
                    
                    'CRYSTALREPORT1.Formulas(5) = "ICU = '" & "Yes" & "'"
                     'SELECT FOR ORGANISM ISOLATED
                     
                    '***** Shraddha Pandit 30/06/09 *****
'                    If Len(Trim(txt_organ1.Text)) > 0 And Len(Trim(TXT_ORGAN2.Text)) = 0 And Len(Trim(TXT_ORGAN3.Text)) = 0 Then
'                        CrystalReport1.Formulas(5) = "ORGAN1 = '" & "Yes" & "'"
'                    ElseIf Len(Trim(txt_organ1.Text)) > 0 And Len(Trim(TXT_ORGAN2.Text)) > 0 And Len(Trim(TXT_ORGAN3.Text)) = 0 Then
'                        CrystalReport1.Formulas(5) = "ORGAN2 = '" & "Yes" & "'"
'                    ElseIf Len(Trim(txt_organ1.Text)) > 0 And Len(Trim(TXT_ORGAN2.Text)) > 0 And Len(Trim(TXT_ORGAN3.Text)) > 0 Then
'                        CrystalReport1.Formulas(5) = "ORGAN3 = '" & "Yes" & "'"
'                    End If
                    
                    If Len(Trim(cmb_org1.Text)) > 0 And Len(Trim(cmb_org2.Text)) = 0 And Len(Trim(cmb_org3.Text)) = 0 Then
                        CRYSTALREPORT1.Formulas(5) = "ORGAN1 = '" & "Yes" & "'"
                    ElseIf Len(Trim(cmb_org1.Text)) > 0 And Len(Trim(cmb_org2.Text)) > 0 And Len(Trim(cmb_org3.Text)) = 0 Then
                        CRYSTALREPORT1.Formulas(5) = "ORGAN2 = '" & "Yes" & "'"
                    ElseIf Len(Trim(cmb_org1.Text)) > 0 And Len(Trim(cmb_org2.Text)) > 0 And Len(Trim(cmb_org3.Text)) > 0 Then
                        CRYSTALREPORT1.Formulas(5) = "ORGAN3 = '" & "Yes" & "'"
                    End If
                    '************************************
                    Set TEMP_RS = rdb.OpenResultset("SELECT * FROM DISLIB.DMULPROCE WHERE  REQNO = '" & mod_reqno & "'")
                    If TEMP_RS.EOF = False Then
                        If Trim(UCase(TEMP_RS("PRO_CODE"))) = "G204" Or Trim(UCase(TEMP_RS("PRO_CODE"))) = "G205" Or Trim(UCase(TEMP_RS("PRO_CODE"))) = "G206" Then
                        CRYSTALREPORT1.Formulas(6) = "TXTPRIMARY_STAIN = ''"
                        Else
                        CRYSTALREPORT1.Formulas(6) = "TXTPRIMARY_STAIN = 'Primary Smear (Gram Stain) '"
                        End If
                    End If
                    'CRYSTALREPORT1.Formulas(13) = "bedno='" & txt_bedno.Text & "'" '17/11/08
                Else
                    'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & mod_reqno & "'"
                    'Commented and added by Shraddha on 02/09/10
                    CRYSTALREPORT1.SelectionFormula = "{PATHONOLF.REQNO} = '" & mod_reqno & "'"
                    CRYSTALREPORT1.ReportFileName = "c:\DISREPORTSECLF\patho_no_report.rpt"
                    CRYSTALREPORT1.Formulas(1) = "proFIN = 'MICROBIOLOGY REPORT'"
                    CRYSTALREPORT1.Formulas(2) = "ENT = '" & "Entered By   " & Trim(cmb_enteryby.Text) & "'"
                    CRYSTALREPORT1.Formulas(3) = "USRID = '" & str & "'"
                    '*********ADDED BY SHWETA 13-AUG-09
                    'Dim TEMP_RS As rdoResultset
'                    SELSTR = "SELECT * FROM DISLIB.DMULPROCE WHERE  REQNO = '" & mod_reqno & "'"
                    Set TEMP_RS = rdb.OpenResultset("SELECT * FROM DISLIB.DMULPROCE WHERE  REQNO = '" & mod_reqno & "'")
                    If TEMP_RS.EOF = False Then
                        If Trim(UCase(TEMP_RS("PRO_CODE"))) = "G204" Or Trim(UCase(TEMP_RS("PRO_CODE"))) = "G205" Or Trim(UCase(TEMP_RS("PRO_CODE"))) = "G206" Then
                        CRYSTALREPORT1.Formulas(4) = "TXTPRIMARY_STAIN = ''"
                        Else
                        CRYSTALREPORT1.Formulas(4) = "TXTPRIMARY_STAIN = 'Primary Smear (Gram Stain) '"
                        End If
                    End If
                    TEMP_RS.Close
                    '**********************************************
'                    AA = "Primary Smear (Gram Stain) "
                    'CRYSTALREPORT1.Formulas(4) = "patho = '" & "No" & "'"
                    'CRYSTALREPORT1.Formulas(13) = "bedno='" & txt_bedno.Text & "'"
                End If
                ElseIf Trim(txt_modcode.Text) = "CM" Then ' for case no. 5(clinical micro)
                If UCase(Trim(CMB_PROC.Text)) = "PREGNANCY TEST (URINARY ?-HCG)" Then
                    CRYSTALREPORT1.SelectionFormula = "{REPORTOTLF.REQNO} = '" & mod_reqno & "'"
                    'CrystalReport1.ReportFileName = "c:\DISREPORTSEC\Other_report.rpt"
                    'Commented and added by Shraddha on 15/09/10
                    CRYSTALREPORT1.ReportFileName = "c:\DISREPORTSECLF\Other_report.rpt"
                    CRYSTALREPORT1.Formulas(1) = "profin = 'MICROBIOLOGY REPORT'"
                    CRYSTALREPORT1.Formulas(2) = "ENT = '" & "Entered By   " & Trim(cmb_enteryby.Text) & "'"
                    CRYSTALREPORT1.Formulas(3) = "USRID = '" & str & "'"
                    CRYSTALREPORT1.Formulas(13) = "bedno='" & txt_bedno.Text & "'"
                    
                    '***** Shraddha 11/002/11 *****
                    If SEL_DMUL_RS("PRO_CODE") = "G256" Then
                        CRYSTALREPORT1.Formulas(17) = "PComment='Yes'"
                    Else
                        CRYSTALREPORT1.Formulas(17) = "PComment=''"
                    End If
                    '******************************
                    
                End If
                If Trim(CMB_PROC.ListCount) > 1 Then
                    'print both portion
                    CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & mod_reqno & "'"
                    CRYSTALREPORT1.ReportFileName = "c:\DISREPORTSEC\Urine_Report.rpt"
                    CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & mod_reqno & "'"
                    CRYSTALREPORT1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
                    CRYSTALREPORT1.Formulas(1) = "proFIN = 'MICROBIOLOGY REPORT'"
                    CRYSTALREPORT1.Formulas(2) = "CODE1 = 'Urine Routine Examination '"
                    CRYSTALREPORT1.Formulas(3) = "ENT = '" & "Entered By   " & Trim(cmb_enteryby.Text) & "'"
                    CRYSTALREPORT1.Formulas(4) = "USRID = '" & str & "'"
                    CRYSTALREPORT1.Formulas(7) = "pathnumber='" & Trim(txt_microno.Text) & "'"
                    CRYSTALREPORT1.Formulas(8) = "ward='" & txt_ward.Text & " /" & txt_bedno.Text & "'"
                    CRYSTALREPORT1.Formulas(9) = "NAM='" & txt_name.Text & "'"
                    CRYSTALREPORT1.Formulas(10) = "STDOB='" & txt_sex.Text & " /" & txt_age.Text & "'"
                    If Opt_patient.Value Then
                        CRYSTALREPORT1.Formulas(11) = "ST='" & txt_category.Text & " /" & txt_status.Text & "'"
                    Else
                        CRYSTALREPORT1.Formulas(11) = "ST='Staff /" & txt_status.Text & "'"
                    End If
                    CRYSTALREPORT1.Formulas(13) = "bedno='" & txt_bedno.Text & "'"
                    CRYSTALREPORT1.WindowState = crptMaximized
                    CRYSTALREPORT1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                    CRYSTALREPORT1.Action = 1
                    '****************************
                    CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & mod_reqno & "'"
                    CRYSTALREPORT1.ReportFileName = "c:\DISREPORTSEC\Stool_Report.rpt"
                    CRYSTALREPORT1.Formulas(1) = "proFIN = 'MICROBIOLOGY REPORT'"
                    CRYSTALREPORT1.Formulas(2) = "CODE2 = 'Stool Routine Examination '"
                    CRYSTALREPORT1.Formulas(3) = "ENT = '" & "Entered By   " & Trim(cmb_enteryby.Text) & "'"
                    CRYSTALREPORT1.Formulas(4) = "USRID = '" & str & "'"
                    CRYSTALREPORT1.Formulas(13) = "bedno='" & txt_bedno.Text & "'"
                ElseIf UCase(Trim(CMB_PROC.Text)) = "URINE EXAMINATION" Then 'print upper portion
                    CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & mod_reqno & "'"
                    CRYSTALREPORT1.ReportFileName = "c:\DISREPORTSEC\Urine_Report.rpt"
                    CRYSTALREPORT1.Formulas(1) = "proFIN = 'MICROBIOLOGY REPORT'"
                    CRYSTALREPORT1.Formulas(2) = "CODE1 = 'Urine Routine Examination '"
                    CRYSTALREPORT1.Formulas(3) = "ENT = '" & "Entered By   " & Trim(cmb_enteryby.Text) & "'"
                    CRYSTALREPORT1.Formulas(4) = "USRID = '" & str & "'"
                    CRYSTALREPORT1.Formulas(13) = "bedno='" & txt_bedno.Text & "'"
                ElseIf UCase(Trim(CMB_PROC.Text)) = "STOOL EXAMINATION" Or UCase(Trim(CMB_PROC.Text)) = "STOOL FOR OCCULT BLOOD" Then 'print lower portion
                    CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & mod_reqno & "'"
                    CRYSTALREPORT1.ReportFileName = "c:\DISREPORTSEC\Stool_Report.rpt"
                    CRYSTALREPORT1.Formulas(1) = "proFIN = 'MICROBIOLOGY REPORT'"
                    CRYSTALREPORT1.Formulas(2) = "CODE2 = 'Stool Routine Examination '"
                    CRYSTALREPORT1.Formulas(3) = "ENT = '" & "Entered By   " & Trim(cmb_enteryby.Text) & "'"
                    CRYSTALREPORT1.Formulas(4) = "USRID = '" & str & "'"
                    CRYSTALREPORT1.Formulas(13) = "bedno='" & txt_bedno.Text & "'"
                    
                    'If SEL_DMUL_RS("PRO_CODE") = "G251" Then
                    'Commented and added by Shraddha on 04/07/2011
                    If Trim(SEL_DMUL_RS("PRO_CODE")) = "G251" Or Trim(SEL_DMUL_RS("PRO_CODE")) = "G102" Then
                    Dim comment As String
'                    comment = ""
'                    comment = "Stool samples collected during menstrual bleeding, constipation induced bleeding, bleeding hemorrhoids or when rectal medication is used may cause positive results" & vbCrLf
'                    comment = comment & "Certain medications may induce gastroinstestinal bleeding and cause false positive reactions hence should be avoided during and prior to the testing period" & vbCrLf
'                    comment = comment & "Diet containing exogenous peroxidases may induce false positive results" & vbCrLf
'                    comment = comment & "Dosages of Vitamin C more than 250mg per day will cause false negative results"
                    
                        'CrystalReport1.Formulas(17) = "OccultComment='" & comment & "'"
                        CRYSTALREPORT1.Formulas(17) = "OccultComment='Yes'"
                    Else
                        CRYSTALREPORT1.Formulas(17) = "OccultComment=''"
                    End If
                    
                End If
            ElseIf UCase(Trim(txt_modcode.Text)) = "MB" Then 'for mycobacteriology
                If UCase(Trim(CMB_PROC.Text)) = "ADA LEVEL" Then
                    'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & mod_reqno & "'"
                    
                    CRYSTALREPORT1.SelectionFormula = "{REPORTOTLF.REQNO} = '" & mod_reqno & "'"
                    
                    'CrystalReport1.ReportFileName = "c:\DISREPORTSEC\Other_report.rpt"
                    'Commented and added by Shraddha on 15/09/10
                    CRYSTALREPORT1.ReportFileName = "c:\DISREPORTSECLF\Other_report.rpt"
                    CRYSTALREPORT1.Formulas(1) = "profin = 'MICROBIOLOGY REPORT'"
                    CRYSTALREPORT1.Formulas(2) = "ENT = '" & "Entered By   " & Trim(cmb_enteryby.Text) & "'"
                    CRYSTALREPORT1.Formulas(3) = "USRID = '" & str & "'"
                    
                Else
                    
                    REPORTSTR = ""
                    If REPORTFLAG = False Then
                            REPORT_QR(0) = mod_reqno
                            Set REPORT_RS = REPORT_QR.OpenResultset(1, 3)
                            REPORTFLAG = True
                    Else
                            REPORT_QR(0) = mod_reqno
                            REPORT_RS.MoreResults
                            REPORT_RS.Requery
                    End If
                    If Len(Trim(REPORT_RS("report"))) = 0 Then
                    REPORTSTR = "No"
                    Else
                    REPORTSTR = "Yes"
                    End If
                    
                    'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & mod_reqno & "' OR {DREQUISITN.REQNO} = '" & mod_reqno & "'"
                    'Commented and added by Shraddha on 01/09/10
                    CRYSTALREPORT1.SelectionFormula = "{MYCOLF.REQNNO} = '" & mod_reqno & "'"
                    CRYSTALREPORT1.ReportFileName = "c:\DISREPORTSECLF\Myco.rpt"
                    ''''
                    CRYSTALREPORT1.Formulas(1) = "proFIN = 'MICROBIOLOGY REPORT'"
                    CRYSTALREPORT1.Formulas(2) = "ENT = '" & "Entered By   " & Trim(cmb_enteryby.Text) & "'"
                    CRYSTALREPORT1.Formulas(3) = "USRID = '" & str & "'"
                    CRYSTALREPORT1.Formulas(13) = "BEDNO ='" & txt_bedno.Text & "'"
                    CRYSTALREPORT1.Formulas(14) = "REPORT1='" & REPORTSTR & "'"
                    ''
                    'CrystalReport1.Formulas(4) = "MB = 'NO'"
                End If
        Else   'for other case
                
                
                CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & mod_reqno & "'"
                
                '***** Shraddha Pandit 20/07/09 *****
           '     If Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1)) = "SEROLOGY" Then
                    SEL_DMUL_QR(0) = mod_reqno
                    Set SEL_DMUL_RS = SEL_DMUL_QR.OpenResultset(1, 3)
                    'If SEL_DMUL_RS("PRO_CODE") = "G145" Or SEL_DMUL_RS("PRO_CODE") = "G146" Or SEL_DMUL_RS("PRO_CODE") = "G148" Or SEL_DMUL_RS("PRO_CODE") = "G149" Or SEL_DMUL_RS("PRO_CODE") = "G113" Then
                    'above statement commented next added by Shraddha on 21/07/09
                    'If SEL_DMUL_RS("PRO_CODE") = "G146" Or SEL_DMUL_RS("PRO_CODE") = "G149" Or SEL_DMUL_RS("PRO_CODE") = "G113" Or SEL_DMUL_RS("PRO_CODE") = "G148" Or SEL_DMUL_RS("PRO_CODE") = "G145" Then
                    'Added for G135 on 07/09/09
                    'If SEL_DMUL_RS("PRO_CODE") = "G146" Or SEL_DMUL_RS("PRO_CODE") = "G149" Or SEL_DMUL_RS("PRO_CODE") = "G113" Or SEL_DMUL_RS("PRO_CODE") = "G148" Or SEL_DMUL_RS("PRO_CODE") = "G145" Or SEL_DMUL_RS("PRO_CODE") = "G135" Then
                    'If SEL_DMUL_RS("PRO_CODE") = "G146" Or SEL_DMUL_RS("PRO_CODE") = "G149" Or SEL_DMUL_RS("PRO_CODE") = "G113" Or SEL_DMUL_RS("PRO_CODE") = "G135" Or SEL_DMUL_RS("PRO_CODE") = "G148". Or SEL_DMUL_RS("PRO_CODE") = "G145" Then
                    'Commented and added by Shraddha on 31/05/2011
                    If SEL_DMUL_RS("PRO_CODE") = "G135" Or SEL_DMUL_RS("PRO_CODE") = "G113" Or SEL_DMUL_RS("PRO_CODE") = "G135" Or SEL_DMUL_RS("PRO_CODE") = "G148" Or SEL_DMUL_RS("PRO_CODE") = "G145" Or SEL_DMUL_RS("PRO_CODE") = "G146" Or SEL_DMUL_RS("PRO_CODE") = "G404" Then
                        'No img needed TRUE
                        'no img needed
                        CRYSTALREPORT1.ReportFileName = "c:\DISREPORTSEC\ report_nologo.rpt"
                    Else
                        'Image needed False
                        'CrystalReport1.ReportFileName = "c:\DISREPORTSEC\Other_report.rpt"
                        'Commented and added by Shraddha on 15/09/10
                        CRYSTALREPORT1.ReportFileName = "c:\DISREPORTSECLF\Other_report.rpt"
                        CRYSTALREPORT1.SelectionFormula = "{REPORTOTLF.REQNO} = '" & mod_reqno & "'"
                        'Selection formula added by Shraddha on 05/10/10
                    End If
            '    Else
                '***********************
              '      CrystalReport1.ReportFileName = "c:\DISREPORTSEC\Other_report.rpt"
             '   End If
                CRYSTALREPORT1.Formulas(1) = "profin = 'MICROBIOLOGY REPORT'"
                CRYSTALREPORT1.Formulas(2) = "ENT = '" & "Entered By   " & Trim(cmb_enteryby.Text) & "'"
                CRYSTALREPORT1.Formulas(3) = "USRID = '" & str & "'"
                CRYSTALREPORT1.Formulas(13) = "bedno='" & txt_bedno.Text & "'"
        End If
   
        CRYSTALREPORT1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
        CRYSTALREPORT1.Formulas(6) = "pathnumber='" & txt_microno.Text & "'" '17/11/08 7->6

'Commented by Shraddha on 18/11/10
'   Else
'    'If ADDENDUMFALG = False Then
'        CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & mod_reqno & "'"
'       CrystalReport1.ReportFileName = "c:\DISREPORTSEC\DISMBADD.rpt" 'SHWETA 03-11-09
'        'CrystalReport1.ReportFileName = "c:\DISREPORTSEC\Other_report.rpt"
'       ' CrystalReport1.ReportFileName = "c:\DISREPORTSEC\Urine_Report.rpt"
'        CrystalReport1.Formulas(1) = "profin = 'MICROBIOLOGY REPORT'"
'        CrystalReport1.Formulas(2) = "ENT = '" & "Entered By   " & Trim(cmb_enteryby.Text) & "'"
'        CrystalReport1.Formulas(3) = "USRID = '" & str & "'"
'        CrystalReport1.Formulas(13) = "bedno='" & txt_bedno.Text & "'"
'        CrystalReport1.Formulas(0) = "REQNO = '" & mod_reqno & "'"
'        CrystalReport1.Formulas(7) = "pathnumber='" & txt_microno.Text & "'"
'        CrystalReport1.Formulas(13) = "bedno='" & txt_bedno.Text & "'"
'   End If
    
    If Not Trim(txt_ward.Text) = "" Then
         If InStr(1, txt_ward.Text, "WARD") = 0 Then
             CRYSTALREPORT1.Formulas(7) = "ward='" & txt_ward.Text & " /" & txt_bedno.Text & "'"
         Else
             CRYSTALREPORT1.Formulas(7) = "ward='" & Trim(Mid(txt_ward.Text, 1, InStr(1, txt_ward.Text, "WARD") - 1)) & "'"
         End If
    Else
            CRYSTALREPORT1.Formulas(7) = "ward="
    End If
     CRYSTALREPORT1.Formulas(8) = "NAM='" & txt_name.Text & "'"
     CRYSTALREPORT1.Formulas(9) = "STDOB='" & txt_sex.Text & " /" & txt_age.Text & "'"
     If Opt_patient.Value Then
        CRYSTALREPORT1.Formulas(10) = "ST='" & txt_category.Text & " /" & txt_status.Text & "'"
     Else
        CRYSTALREPORT1.Formulas(10) = "ST='Staff /" & txt_status.Text & "'"
     End If
     
   Else
    MsgBox "Please select valid Requisition No.", vbOKOnly
    Exit Sub
 End If

    '********** Shraddha Pandit 07/09/09 **********
    'to windows if from query status else to Printer
    If QrySts = True Then
        CRYSTALREPORT1.Destination = crptToWindow
        QrySts = False
    Else
        'CrystalReport1.Destination = crptToPrinter 'SHWETA PRINTER
        CRYSTALREPORT1.Destination = crptToWindow
    End If
    '**********************************************
    
    'CrystalReport1.Destination = crptToPrinter 'added on 05/06/09 by Shraddha 'COMMENTED BY SHWETA
    'CrystalReport1.Destination = crptToWindow 'SHWETA 13-AUG-09
    CRYSTALREPORT1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
    'CrystalReport1.WindowState = crptMaximized 'Commented on 05/06/09 by Shraddha
    CRYSTALREPORT1.Action = 1
End If
On Error GoTo AA
TEMP_RS.Close
AA:
    Exit Sub
MSG:
    MsgBox ERROR(Err)
End Sub

Private Sub cmd_save_Click()
   '----mrunal
    str = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USR_ID='A" & usr_name & "' "
    Set REPWTRN_QR = rdb.CreateQuery("", str)
   Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)
    If REPWTRNRS.RowCount > 0 Then
    cmb_final.Text = REPWTRNRS("USERNAME")
    End If
    '----
 Label27.Caption = ""
If Report_microbiology1.Text2.Text <> "" Then
If Report_microbiology1.Text2.Text = "2" Then
cmb_reqno.Text = req
Else
   cmb_reqno.Text = Report_microbiology1.Text2.Text
 End If
  End If
'modified by mrunal--25_2
    I = InStr(cmb_modality.Text, "$")
    txt_modcode.Text = Mid(cmb_modality.Text, I + 1, 2)
  
    If txt_modcode.Text = "BA" Then
    
If cmb_patho = "Yes" Then
If PRO_CODE <> "G111" And PRO_CODE <> "G209" And PRO_CODE <> "G206" Then
 If cmb_org1.Visible = True Or cmb_org2.Visible = True Or cmb_org3.Visible = True Then
    If cmb_org1 = "" Then
        If cmb_org2 = "" Then
            If cmb_org3 = "" Then
             MsgBox "Atleast 1 organisms is needed", vbOKOnly
            
                    cmb_org1.SetFocus
                 
                    Exit Sub
                       End If
                    End If
                    End If
            End If
        End If
    End If
End If
'end
'ADDED BY MILIND ON 16-10-06 FOR SEPARATING REQNO FROM LABNO.
Dim a As Integer
a = InStr(1, cmb_reqno, "$")
mod_reqno = req_chglen(Mid(Trim(cmb_reqno), 1, a - 1))
 UPDATERTIMELOG1(0) = Trim(txt_spm.Text)
    UPDATERTIMELOG1(1) = mod_reqno

    UPDATERTIMELOG1.Execute
'UPTO HERE

If new_flag Then    'if new record then use insert query
'*******  Commented on 05/06/09 by Shraddha *******
'    If Len(Trim(cmb_enteryby)) = 0 Then
'        MsgBox "The Enter by field is EMPTY.", vbCritical
'        cmb_enteryby.SetFocus
'        Exit Sub
'    End If
'
'    If Len(Trim(cmb_final)) = 0 Then
'        MsgBox "The Finalized by field is EMPTY.", vbCritical
'        cmb_final.SetFocus
'        Exit Sub
'    End If
'
'    If Len(Trim(txt_c_report.Text)) >= 200 Then
'         MsgBox "The Report Description Max. length is 200 Characters Only"
'         txt_c_report.SetFocus
'         Exit Sub
'    End If
'************ Till here 05/06/09 Shraddha *************
    While InStr(txt_addendum, Chr(9)) > 0
            
        Dim beforst As String, afterst As String
        Dim K As Long
        K = InStr(txt_addendum, Chr(9))
        beforst = Mid(txt_addendum, 1, K - 1) & "      "
        afterst = Mid(txt_addendum, K + 1)
        txt_addendum = beforst & afterst
            
    Wend
        
    While InStr(txt_addendum, Chr(146)) > 0

        K = InStr(txt_addendum, Chr(146))
        beforst = Mid(txt_addendum, 1, K - 1) & "`"
        afterst = Mid(txt_addendum, K + 1)
        txt_addendum = beforst & afterst
       
    Wend
        
    '********* Shraddha Pandit 24/10/08 ************

    If new_flag Then
        If Frame2.Visible Then
            If txt_test.Text = "" Then
                MsgBox "Please enter Test Performed field"
                Exit Sub
            End If
            If txt_testresult.Text = "" Then
                MsgBox "Please enter Result field"
                Exit Sub
            End If
                        
            str = "insert into DISLIB.micro_test(REQNO,DRUG,MIC,UOM,RESULT,TEST,TEST_RES,T_DATE,T_TIME,U_ID)values(?,?,?,?,?,?,?,?,?,?)"
            Set GINS_QR = rdb.CreateQuery("", str)
    
            Dim I2 As Integer
            I2 = 1
            While I2 <= countr - 1
        
                GINS_QR(0) = RQN
                GINS_QR(1) = MSFlexGrid1.TextMatrix(I2, 0)
                GINS_QR(2) = MSFlexGrid1.TextMatrix(I2, 1)
                GINS_QR(3) = MSFlexGrid1.TextMatrix(I2, 2)
                GINS_QR(4) = MSFlexGrid1.TextMatrix(I2, 3)
                GINS_QR(5) = txt_test.Text
                GINS_QR(6) = txt_testresult
                GINS_QR(7) = Format(txt_date.Text, "YYYY-MM-DD")
                'GINS_QR(8) = Format(txt_time.Text, "HH:MM:SS")
                GINS_QR(8) = Format(time1, "HH:MM:SS")
                GINS_QR(9) = UID
            
                GINS_QR.Execute

                I2 = I2 + 1
            Wend
       
            '***** Shraddha 04/07/2011 *****
            str = "insert into DISLIB.dmycobacte (reqno,comment) values(?,?)"
            Set ins_mccomment_qr = rdb.CreateQuery("", str)
            
            ins_mccomment_qr(0) = RQN
            ins_mccomment_qr(1) = Trim(txt_MCcomment.Text)
            ins_mccomment_qr.Execute
            '*******************************
       
            str = "update rislib.rrequisitn set reqnstat='T',enterby=?,preportby=?,preportdt=? where caseno=? and reqnno=?"
            Set UPDReqRQ = rdb.CreateQuery("", str)
            UPDReqRQ(0) = cmb_enteryby.Text
            UPDReqRQ(1) = cmb_final.Text 'cmb_final.Text MODIFIED BY MRUNAL 23_2
            UPDReqRQ(2) = Format(txt_date.Text, "YYYY-MM-DD")
            UPDReqRQ(3) = MaskEdBox1.FormattedText
            UPDReqRQ(4) = mod_reqno
            UPDReqRQ.Execute
       
            'MsgBox "Record successfully saved"
            Label27.Caption = "Record Has Been Saved Successfully"
        End If
        '******************************************
        If Frame3.Visible = True Then '11/11/08
        
            If txt_btest.Text = "" Then
                'MsgBox "Please enter Test Performed field"
               ' Exit Sub
            End If
                       
            If txt_bresult.Text = "" Then
               'MsgBox "Please enter Test Performed field"
               ' Exit Sub
            End If
            
            str = "insert into DISLIB.dmoprismer(REQNO,OFINDING,REPORT)values(?,?,?)"
            Set GBINS_QR = rdb.CreateQuery("", str)
    
            GBINS_QR(0) = RQN
            GBINS_QR(1) = txt_btest
            GBINS_QR(2) = txt_bresult
                        
            GBINS_QR.Execute
            
            str = "update rislib.rrequisitn set reqnstat='T',enterby=?,preportby=?,preportdt=? where caseno=? and reqnno=?"
            Set UPDReqBRQ = rdb.CreateQuery("", str)
            UPDReqBRQ(0) = cmb_enteryby.Text
            UPDReqBRQ(1) = cmb_final.Text 'cmb_final.Text MODIFIED BY MRUNAL 23_2
            UPDReqBRQ(2) = Format(txt_date.Text, "YYYY-MM-DD")
            UPDReqBRQ(3) = MaskEdBox1.FormattedText
            UPDReqBRQ(4) = mod_reqno
            UPDReqBRQ.Execute
       
            'MsgBox "Record successfully saved"
             Label27.Caption = "Record Has Been Saved Successfully"
              cmd_commit.Visible = True
        End If
      
    End If
    
'    new_flag = False
'    mod_flag = True
    '***********************************************
    btnenabl "SAVE"
    new_flag = False
'        If (frm_bacteria.Visible) Then
'             If Trim(UCase(cmb_patho.Text)) = "YES" Then
'                    If grid_dmbanti.Rows = 1 Then
'                MsgBox "errOR"
'                    End If
'             End If
'        End If
        
    'check which test is performed
    If (frm_bacteria.Visible) And (Frame3.Visible = False) Then 'Bacteriology 13/11/08 Shraddha
        'insert in DMSPONANIB AND DMPRISMER
        insert_bacteriology
         ' MsgBox "Record Saved Successfully!!!!!"
        'cmd_commit_Click 'Shraddha 17/11/08
        'cmd_print_Click 'Shraddha 17/11/08
        'cmd_cancel_Click 'Shraddha 17/11/08
            
'        If DEPT_CODE = "GZZ" Then
'            cmd_new_Click
'            cmb_modality.Text = modtxt
'
'+xt = modtxt
'            MaskEdBox1.SetFocus
'        End If
'        Exit Sub
    ElseIf Frame3.Visible = True Then
'        cmd_commit_Click
'        cmd_print_Click
'        cmd_cancel_Click
'        If DEPT_CODE = "GZZ" Then
'            cmd_new_Click
'            cmb_modality.Text = modtxt
'            txt_modcode.Text = modtxt
'            MaskEdBox1.SetFocus
'        End If
       ' Exit Sub
    ElseIf frm_clinic.Visible Then 'for Clinical Microbiology
            'now in this check for stool or urine
            
            If frm_urineexam.Enabled And frm_stoolexam.Enabled Then
                insert_stool
                insert_urine
                'MsgBox "Record successfully saved" '(insert) commented on 05/06/09 by Shraddha
                Label27.Caption = "Record Has Been Saved Successfully"
               ' Exit Sub
            ElseIf frm_stoolexam.Enabled Then
                'insert in stool table  DMSRE
                insert_stool
               ' MsgBox "Record successfully saved"
                Label27.Caption = "Record Has Been Saved Successfully"
                  cmd_commit.Visible = True
                Exit Sub
            ElseIf frm_urineexam.Enabled Then
                insert_urine
              '  MsgBox "Record successfully saved"
               Label27.Caption = "Record Has Been Saved Successfully"
                  cmd_commit.Visible = True
                Exit Sub
            End If
    ElseIf frm_myco.Visible Then
        'insert into table dmycobacte
        insert_myco
        'MsgBox "Record successfully saved" '(insert) Commented on 08/06/09 by Shraddha
         Label27.Caption = "Record Has Been Saved Successfully"
        btnenabl "SAVE"
       ' Exit Sub
     ElseIf frmSerology.Visible = True Then
        insert_GZZSEROLOG
        'Exit Sub
    ElseIf frm_others.Visible Then
        'insert in free text table DMBSEMCMO
        insert_other
       ' MsgBox "Record successfully saved" 'Insert
        Label27.Caption = "Record Has Been Saved Successfully"
       ' Exit Sub
    End If
        
    If Len(Trim(txt_addendum.Text)) > 600 Then
        MsgBox "Addendum Max. of 600 Characters Only.", vbInformation
        txt_addendum.SetFocus
        Exit Sub
    End If
End If
    'cmd_commit.Enabled = True
    If mod_flag Then 'else if modifying a record then use update query
        ins_mod_var = "mod"
        If Len(Trim(cmb_enteryby)) = 0 Then
            MsgBox "The Enter by field is EMPTY.", vbCritical
            cmb_enteryby.SetFocus
        Exit Sub
        End If
    
'        If Len(Trim(cmb_final)) = 0 Then
'            MsgBox "The Finalized by field is EMPTY.", vbCritical
'            cmb_final.SetFocus
'            Exit Sub
'        End If
    
        '    If Len(Trim(txt_c_report.Text)) >= 200 Then
        '         MsgBox "The Report Description Max. length is 200 Characters Only"
        '         txt_c_report.SetFocus
        '         Exit Sub
        '    End If
        
            '***************** 30/10/08 Shraddha Pandit *********
    If Frame2.Visible = True Then
    
        If txt_test.Text = "" Then
            MsgBox "Please enter Test Performed field"
            Exit Sub
        End If
        If txt_testresult.Text = "" Then
            MsgBox "Please enter Result field"
            Exit Sub
        End If
        
        str = "update dislib.micro_test set DRUG=?,MIC=?,UOM=?,RESULT=?,TEST=?,TEST_RES=?,T_DATE=?,T_TIME=?,U_ID=? where reqno=? and drug=?"
        Set GUPD_QR = rdb.CreateQuery("", str)

        Dim I3 As Integer
        I3 = 1
        'While I3 <= countr - 1
        While I3 <= countr
            GUPD_QR(0) = MSFlexGrid1.TextMatrix(I3, 0)
            GUPD_QR(1) = MSFlexGrid1.TextMatrix(I3, 1)
            GUPD_QR(2) = MSFlexGrid1.TextMatrix(I3, 2)
            GUPD_QR(3) = MSFlexGrid1.TextMatrix(I3, 3)
            GUPD_QR(4) = txt_test.Text
            GUPD_QR(5) = txt_testresult.Text
            GUPD_QR(6) = Format(txt_date.Text, "YYYY-MM-DD")
            'GUPD_QR(7) = Format(txt_time.Text, "HH:MM:SS")
            GUPD_QR(7) = Format(time1, "HH:MM:SS")
            GUPD_QR(8) = UID
            GUPD_QR(9) = RQN
            GUPD_QR(10) = MSFlexGrid1.TextMatrix(I3, 0)
            GUPD_QR.Execute

            I3 = I3 + 1
        Wend
       
        '***** Shraddha 04/07/2011 *****
        str = "update DISLIB.dmycobacte set comment=? where reqno=?"
        Set upd_mccomment_qr = rdb.CreateQuery("", str)
        
        upd_mccomment_qr(0) = Trim(txt_MCcomment.Text)
        upd_mccomment_qr(1) = RQN
        upd_mccomment_qr.Execute
        '*******************************
        str = "update rislib.rrequisitn set reqnstat='T',enterby=?,preportby=?,preportdt=? where caseno=? and reqnno=?"
        Set UPDReqRQ = Nothing
        Set UPDReqRQ = rdb.CreateQuery("", str)
        UPDReqRQ(0) = cmb_enteryby.Text
        UPDReqRQ(1) = cmb_final.Text 'cmb_final.Text modified by mrunal 23_2
        UPDReqRQ(2) = Format(txt_date.Text, "YYYY-MM-DD")
        UPDReqRQ(3) = MaskEdBox1.FormattedText
        UPDReqRQ(4) = mod_reqno
        UPDReqRQ.Execute
       
        '******** Shraddha Pandit 04/11/08*******
        'Display msgbox only if others frame is invisible otherwise it will be displayed twice
        
        If frm_others.Visible = False Then
           ' MsgBox "Record successfully modified"
            Label27.Caption = "Record Has Been modified Successfully"
              cmd_commit.Visible = True
        End If
    End If
        '*******************************************
        '------- Shraddha Pandit 12/11/08 ----------
        
        If Frame3.Visible = True Then
        
            If txt_btest.Text = "" Then
                MsgBox "Please enter Test Performed field"
                Exit Sub
            End If
                       
            If txt_bresult.Text = "" Then
                MsgBox "Please enter Test Performed field"
                Exit Sub
            End If
        str = "update dislib.dmoprismer set OFINDING=?,REPORT=? where reqno=?"
        Set GBUPD_QR = rdb.CreateQuery("", str)

        GBUPD_QR(0) = txt_btest.Text
        GBUPD_QR(1) = txt_bresult.Text
        GBUPD_QR(2) = RQN
                
        GBUPD_QR.Execute
              
        str = "update rislib.rrequisitn set reqnstat='T',enterby=?,preportby=?,preportdt=? where caseno=? and reqnno=?"
        Set UPDReqBRQ = Nothing
        Set UPDReqBRQ = rdb.CreateQuery("", str)
        UPDReqBRQ(0) = cmb_enteryby.Text
        UPDReqBRQ(1) = cmb_final.Text 'cmb_final.Text modified by mrunal 23_2
        UPDReqBRQ(2) = Format(txt_date.Text, "YYYY-MM-DD")
        UPDReqBRQ(3) = MaskEdBox1.FormattedText
        UPDReqBRQ(4) = mod_reqno
        UPDReqBRQ.Execute
       
'        ******** Shraddha Pandit 04/11/08*******
'        Display msgbox only if others frame is invisible otherwise it will be displayed twice

        If frm_others.Visible = False Then
            'MsgBox "Record successfully modified"
             Label27.Caption = "Record Has Been modified Successfully"
              cmd_commit.Visible = True
        End If
    End If
    
        '--------------------------------------------
        
    mod_flag = False
    btnenabl "SAVE"
    'check which test is performed
    txt_c_1.Visible = False
    Dim str2 As String
    str2 = txt_c_1.Text
    If frm_bacteria.Visible Then 'for Bacteriology
        'update in DMSPONANIB AND DMPRISMER
        'before update get result from database server and assign them to fields by calling getresult
        upd_1_qr(0) = txt_epith.Text
        upd_1_qr(1) = txt_pus.Text
        upd_1_qr(2) = txt_other.Text
        'upd_1_qr(3) = txt_ziehl.Text
        upd_1_qr(3) = txt_ziehl.Text
        upd_1_qr(4) = txt_koh.Text
        upd_1_qr(5) = cmb_gpc.Text
        upd_1_qr(6) = cmb_gnc.Text
        upd_1_qr(7) = cmb_gpb.Text
        upd_1_qr(8) = cmb_gnb.Text
        upd_1_qr(9) = cmb_india.Text
        If Len(txt_comment.Text) > 0 Then
            upd_1_qr(14) = Null
            upd_1_qr(14) = txt_comment.Text
        Else
            upd_1_qr(14) = Null
            upd_1_qr(14) = " "
        End If
        upd_1_qr(15) = mod_reqno
        'STR = "update dislib.dmoprismer set 0epithecell=? ,1puscell=?,2ofinding=?,3znstain=?,4kohamt=?,5gpc=?,6gnc=?,7gpb=?,8gnb=?,9iipc=?,10orgso1=?,11orgso2=?,12orgso3=?,13REPORT=?,14comments=? where 15reqno=? "
        If cmb_patho.Text = "No" Then     'patho
            upd_1_qr(10) = Null
            upd_1_qr(11) = Null
            upd_1_qr(12) = Null
            
            If Len(txt_report.Text) > 0 Then
                upd_1_qr(13) = Null
                upd_1_qr(13) = txt_report.Text
            Else
                upd_1_qr(13) = Null
                upd_1_qr(13) = " "
            End If
            
            upd_1_qr.Execute
        Else
            '******** Commented and added on 30/06/09 by Shraddha **********
'            If Len(txt_organ1.Text) > 0 Then
'                upd_1_qr(10) = Null
'                upd_1_qr(10) = txt_organ1.Text
'            Else
'                upd_1_qr(10) = Null
'            End If
            
            
            If Len(cmb_org1.Text) > 0 Then
                upd_1_qr(10) = Null
                upd_1_qr(10) = cmb_org1.Text
            Else
                upd_1_qr(10) = Null
            End If
                  
            
'            If Len(TXT_ORGAN2.Text) > 0 Then
'                upd_1_qr(11) = Null
'                upd_1_qr(11) = TXT_ORGAN2.Text
'            Else
'                upd_1_qr(11) = Null
'            End If
            
            If Len(cmb_org2.Text) > 0 Then
                upd_1_qr(11) = Null
                upd_1_qr(11) = cmb_org2.Text
            Else
                upd_1_qr(11) = Null
            End If
            
'            If Len(TXT_ORGAN3.Text) > 0 Then
'                upd_1_qr(12) = Null
'                upd_1_qr(12) = TXT_ORGAN3.Text
'            Else
'                upd_1_qr(12) = Null
'            End If

            If Len(cmb_org3.Text) > 0 Then
                upd_1_qr(12) = Null
                upd_1_qr(12) = cmb_org3.Text
            Else
                upd_1_qr(12) = Null
            End If
            
            '***************************************************************
            
            upd_1_qr(13) = Null
            upd_1_qr.Execute
                    
'''Done By Kanchan
            UPDATE_DMBAANTI
        End If 'of patho
        update_rislib
         Label27.Caption = "Record Has Been modified Successfully"
       ' MsgBox "Record Successfully modified!", vbOKOnly 'Commented on 05/06/09 by Shrad
        Exit Sub
    ElseIf frm_clinic.Visible Then 'for Clinical Microbiology
        'now in this check for stool or urine
            If frm_urineexam.Enabled And frm_stoolexam.Enabled Then
                UPDATE_URINE
                UPDATE_STOOL
                update_rislib
               ' MsgBox "Record Updated Successfully", vbOKOnly
                Label27.Caption = "Record Has Been modified Successfully"
                  cmd_commit.Visible = True
                Exit Sub
            ElseIf frm_urineexam.Enabled Then
                'update in urine table DMURE
                   UPDATE_URINE
                   update_rislib
                   'MsgBox "Record Updated Successfully", vbOKOnly
                    Label27.Caption = "Record Has Been modified Successfully"
                     cmd_commit.Visible = True
                   Exit Sub
            ElseIf frm_stoolexam.Enabled Then
                'update in stool table  DMSRE
                    UPDATE_STOOL
                    update_rislib
                 '   MsgBox "Record Updated Successfully", vbOKOnly
                  Label27.Caption = "Record Has Been modified Successfully"
                    Exit Sub
            End If
    ElseIf frm_others.Visible Then
        If SELDMBFLAG = False Then
            'SELDMBSEMCOM_QR(0) = Trim(txt_test_perform.Text)
            SELDMBSEMCOM_QR(0) = mod_reqno
            Set SELDMBSEMCOM_RS = SELDMBSEMCOM_QR.OpenResultset(1, 3)
            SELDMBFLAG = True
        Else
            'SELDMBSEMCOM_QR(0) = Trim(txt_test_perform.Text)
            SELDMBSEMCOM_QR(0) = mod_reqno
            SELDMBSEMCOM_RS.MoreResults
            SELDMBSEMCOM_RS.Requery
        End If
        If SELDMBSEMCOM_RS.RowCount > 0 Then
            
                'update in free text table DMBSEMCMO
                'If Len(txt_refer_service.Text) > 0 Then
                    upd_0_qr(0) = Null
                '   upd_0_qr(0) = txt_refer_service.Text
                'Else
                '   upd_0_qr(0) = Null
                '   upd_0_qr(0) = " "
                ' End If
            
            If Len(txt_test_perform.Text) > 0 Then
                upd_0_qr(1) = Null
                upd_0_qr(1) = Trim(txt_test_perform.Text)
            Else
                upd_0_qr(1) = Null
                upd_0_qr(1) = " "
            End If
            If Len(txt_result.Text) > 0 Then
                upd_0_qr(2) = Null
                upd_0_qr(2) = txt_result.Text
            Else
                upd_0_qr(2) = Null
                upd_0_qr(2) = " "
            End If
            upd_0_qr(3) = mod_reqno
            upd_0_qr.Execute
            update_rislib
            'MsgBox "Record modified Successfully", vbOKOnly '(update) commented on 08/06/09 by Shraddha
            Exit Sub
        Else
            insert_other
        End If
    ElseIf frmSerology.Visible = True Then
        str = "select reqno from dislib.GZZSEROLOG where reqno='" & PRO_CODE & "'"
        Set RS = rdb.OpenResultset(str, 1, 1)
        If RS.EOF Then
            update_GZZSEROLOG
        Else
            insert_GZZSEROLOG
        End If
        RS.Close
    ElseIf frm_myco.Visible Then
        
        If CHKFLAG = False Then
            CHECKREC_QR(0) = mod_reqno
            Set CHECKREC_RS = CHECKREC_QR.OpenResultset(1, 3)
            CHECKFLAG = True
        Else
            CHECKREC_QR(0) = mod_reqno
            CHECKREC_RS.MoreResults
            CHECKREC_RS.Requery
        End If
        
        If CHECKREC_RS.RowCount = 0 Then
            insert_myco
        Else
            'update in free text table Dmycobacte
            upd_7_qr(0) = cmb_stain.Text
            upd_7_qr(1) = cmb_organisms.Text
            If Len(txt_c_report.Text) > 0 Then
                upd_7_qr(2) = Null
                upd_7_qr(2) = txt_c_report.Text
            Else
                upd_7_qr(2) = Null
                upd_7_qr(2) = " "
            End If
            If Len(txt_c_comment.Text) > 0 Then
                upd_7_qr(3) = Null
                upd_7_qr(3) = txt_c_comment.Text
            Else
                upd_7_qr(3) = Null
                upd_7_qr(3) = " "
            End If
            upd_7_qr(4) = mod_reqno
            upd_7_qr.Execute
        End If
        
'***end if
'**************** FOR UPDATING OF MBDAANTI TABLE VALUES   ***********
            'AMIT CODE HERE
        sel_forupd2_qr(0) = mod_reqno
        Set sel_forupd2_rs = sel_forupd2_qr.OpenResultset(1, 3)
        If Not sel_forupd2_rs.EOF Then
            I = grid_c_anti.Rows
            srno = 1
            For I = 1 To grid_c_anti.Rows - 1
                del_dmbaanti_qr1(0) = mod_reqno
                del_dmbaanti_qr1(1) = Trim(grid_c_anti.TextMatrix(srno, 0))
                del_dmbaanti_qr1.Execute
                srno = srno + 1
            Next I
            srno = 1
            For I = 1 To grid_c_anti.Rows - 1
                If Not Trim(grid_c_anti.TextMatrix(srno, 1)) = "" Then
                    ins_2_qr(0) = Trim(mod_reqno)
                    ins_2_qr(1) = Trim(grid_c_anti.TextMatrix(srno, 0))
                    ins_2_qr(2) = grid_c_anti.TextMatrix(srno, 1)
                    ins_2_qr(3) = Null
                    ins_2_qr(4) = Null
                    ins_2_qr.Execute
                End If
                srno = srno + 1
            Next I
        Else
            srno = 1
            For I = 1 To grid_c_anti.Rows - 1
                If Not Trim(grid_c_anti.TextMatrix(srno, 1)) = "" Then
                    ins_2_qr(0) = Trim(mod_reqno)
                    ins_2_qr(1) = Trim(grid_c_anti.TextMatrix(srno, 0))
                    ins_2_qr(2) = grid_c_anti.TextMatrix(srno, 1)
                    ins_2_qr(3) = Null
                    ins_2_qr(4) = Null
                    ins_2_qr.Execute
                End If
                srno = srno + 1
            Next I
                         
        End If
            ''AMIT END
''''            i = 0
''''            srno = 0
''''            i = grid_c_anti.Rows
''''            srno = 2
''''  If Not i = 1 Then 'srno = 1 Then
''''                If Not (Trim(grid_c_anti.TextMatrix(srno, 1)) = "") Then
''''                    While Not srno = i
''''                        upd_2_qr(0) = grid_c_anti.TextMatrix(srno, 1) 'get first value and insert
''''                        upd_2_qr(1) = Null
''''                        upd_2_qr(2) = Null
''''                        upd_2_qr(3) = mod_reqno
''''                        upd_2_qr(4) = Trim(grid_c_anti.TextMatrix(srno, 0)) 'first get name of antibiotic and insert
''''                        upd_2_qr.Execute
''''                        srno = srno + 1
''''                    Wend
''''' antibio1=?,antibio2=?,antibio3=? where reqno=? and antibiotic=?
''''               End If
'''' End If
'**************** FOR UPDATING OF MBDAANTI TABLE VALUES ***********

        update_rislib
            
      '  MsgBox "Record modified Successfully", vbOKOnly
       Label27.Caption = "Record Has Been modified Successfully"
    End If
  cmd_save.Enabled = True
End If
If ADDENDUMFALG = True Then
                ''********SECURITY***************************
    CURUSRDET_QR(0) = Trim(cmb_final.Text)
    CURUSRDET_RS.MoreResults
    CURUSRDET_RS.Requery
            
    If CURUSRDET_RS.RowCount > 0 Then
    
        If LOGINTYP > CURUSRDET_RS("usrtype") Then
                      ''   no action as he can commit
        ElseIf LOGINTYP < CURUSRDET_RS("usrtype") Then
            MsgBox "Sorry! U'r not an Authorized person to write an Addendum !!", vbInformation, "Pathlogy"
            cmd_save.SetFocus
            Exit Sub
        End If
        If CURUSRDET_RS("ADDENDUM") <> "YES" Then
            MsgBox "Sorry! You are not authorised to Write Addendum of this Modality"
            cmd_save.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Sorry! The Security details of this user is not available check up in DISLNKUSR"
        cmd_commit.SetFocus
        Exit Sub
    End If
        ''********SECURITY***************************
                    
    If SELDADDENDUMFLAG = False Then
        SELDADDENDUM_QR(0) = mod_reqno
        Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
        SELDADDENDUMFLAG = True
    Else
        SELDADDENDUM_QR(0) = mod_reqno
        SELDADDENDUM_RS.MoreResults
        SELDADDENDUM_RS.Requery
    End If
            
    If SELDADDENDUM_RS.RowCount > 0 Then
        INSDADDENDUM_QR(0) = mod_reqno
        INSDADDENDUM_QR(1) = SELDADDENDUM_RS("ADDEND_NO") + 1
        INSDADDENDUM_QR(2) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        INSDADDENDUM_QR(3) = time1
        INSDADDENDUM_QR(4) = cmb_addendumby.Text
        INSDADDENDUM_QR(5) = Null
        INSDADDENDUM_QR(5) = txt_addendum.Text
        INSDADDENDUM_QR.Execute
    Else
        INSDADDENDUM_QR(0) = mod_reqno
        INSDADDENDUM_QR(1) = "1"
        INSDADDENDUM_QR(2) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        INSDADDENDUM_QR(3) = time1
        INSDADDENDUM_QR(4) = cmb_addendumby.Text
        INSDADDENDUM_QR(5) = Null
        INSDADDENDUM_QR(5) = txt_addendum.Text
        INSDADDENDUM_QR.Execute
        ' Mdified by mrunal on 21_2_13
           cmd_save.Enabled = True
             ' cmd_save.Enabled = False
           '------end
     
        txt_addendum.Enabled = False
        cmb_addendumby.Enabled = False
                    
     '   MsgBox "Record Updated Successfully!!!!!"
                     Label27.Caption = "Record Has Been modified Successfully"
    End If
End If
             cmd_commit.Visible = True
'once inserted or modified then clear all field
MODE = ""
'End If
  
cmd_save.Enabled = False

Label27.Caption = "Record Saved Sucessfully."
End Sub

Public Sub cmd_view_Click()
   '----mrunal
    str = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USR_ID='A" & usr_name & "' "
    Set REPWTRN_QR = rdb.CreateQuery("", str)
   Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)
    If REPWTRNRS.RowCount > 0 Then
    cmb_final.Text = Trim(REPWTRNRS("USERNAME"))
    End If
    '----
 Label27.Caption = ""
If Report_microbiology1.Text2.Text <> "" Then
cmb_reqno.clear
Else
Dim L1 As Integer
' For L1 = 0 To cmb_final.ListCount
'    If Trim(cmb_enteryby.List(0)) = cmb_final.List(L1) Then
'        cmb_final.Text = cmb_final.List(L1)
'    End If
'    Next L1
    cmb_final.Enabled = False
  cmd_commit.Visible = True
On Error GoTo MSG:

btnenabl ("VIEW")



Set sel_deptrs = sel_deptqr.OpenResultset(1, 3)
    'get the result
    While Not sel_deptrs.EOF
            cmb_dept.AddItem Trim(sel_deptrs("dept_name")) & "$" & Trim(sel_deptrs("dept_code"))
            sel_deptrs.MoveNext
    Wend
    
sel_deptrs.Close
If Trim(cmb_dept) = "" Then
        For I = 0 To (cmb_dept.ListCount - 1)
            If Trim(cmb_dept.List(I)) = "MICROBIOLOGY$GZZ" Then
                cmb_dept.ListIndex = I
            End If
        Next I
End If
'copied and pasted
I = InStr(cmb_dept.Text, "$")
    txt_deptcode.Text = Mid(cmb_dept.Text, I + 1, 3)
    sel_moadlityqr(0) = Trim(txt_deptcode.Text)
    Set sel_modalityrs = sel_moadlityqr.OpenResultset(1, 3)
    If sel_modalityrs.RowCount > 0 Then
        cmb_modality.clear
        While Not sel_modalityrs.EOF
            cmb_modality.AddItem Trim(sel_modalityrs("modlitydes")) & "$" & Trim(sel_modalityrs("modalitycd"))
            sel_modalityrs.MoveNext
        Wend
        sel_modalityrs.Close
    Else
        sel_modalityrs.Close
    End If
    
    '************** Shraddha Pandit 05/06/09 ************
    cmb_modality.Text = cmb_modality.List(0)
    cmb_modality_Click
    '****************************************************
    End If
set_flag "VIEW"


Opt_patient.Value = True
cmd_print.Enabled = True
'cmb_dept.SetFocus


'MaskEdBox1.SetFocus '13-AUG-09 SHWETA


Exit Sub
MSG:
    MsgBox Err.Description
    cmd_commit.Visible = True
End Sub


Private Sub Command1_Click()
'If Len(Trim(MaskEdBox1.Text)) > 0 Then
'    Dim VAR1
'    'VAR = Shell("C:\WINDOWS\Desktop\EMR 04.08.05.EXE" & Trim(Msk_caseno.FormattedText), vbNormalFocus)
'    VAR1 = MaskEdBox1.FormattedText
'    VAR = Shell("C:\EMR1.EXE " & VAR1, vbNormalFocus)
'    AppActivate VAR
'Else
'    MsgBox "Please Enter The Case No."
'End If
If Len(Trim(MaskEdBox1.Text)) > 0 Then
    Dim VAR1
    VAR1 = MaskEdBox1.FormattedText
    Dim strLink As String
    'strLink = "http://intranet.tmc.gov.in/emrweb/default.asp" '10.1.4.9
    strLink = "http://intranet.tmc.gov.in/emrtmhact/TOPBAR.ASP?CASE_NO=" & VAR1 & ""
    Shell "c:\program files\internet explorer\iexplore " & strLink
Else
    MsgBox "Please Enter The Case No."
End If
End Sub

Private Sub Command2_Click()
micro_myco.show
End Sub

Private Sub Form_Load()
NEWF = False
DATETIME
txt_date.Text = Format(DATE11, "DD/MM/YYYY")
txt_time.Text = time1
set_visible "All"
no_req = False

prev_flag = False 'Addded by Shraddha on 11/06/09


''***** Shraddha 31/01/11 *****
'txt_stoolcomments.clear
'txt_stoolcomments.AddItem "Stool samples collected during menstrual bleeding, constipation induced bleeding, bleeding hemorrhoids or when rectal medication is used may cause positive results"
'txt_stoolcomments.AddItem "Certain medications may induce gastroinstestinal bleeding and cause false positive reactions hence should be avoided during and prior to the testing period"
'txt_stoolcomments.AddItem "Diet containing exogenous peroxidases may induce false positive results"
'txt_stoolcomments.AddItem "Dosages of Vitamin C more than 250mg per day will cause false negative results"
''*****************************

'addendum
txt_addendum.Enabled = False
cmb_addendumby.Enabled = False
'initialize all combo with default value
initialize_combo
cmb_india.AddItem "Negative"
cmb_india.AddItem "Positive"
cmb_india.AddItem " "
'txt_comment.MultiLine = True
'txt_comment.ScrollBars = vbBoth
'***************************************************************************************
'                       QUERY DEFINITION
'***************************************************************************************
If MBREPORTFALG = False Then
    MBREPORTFALG = True

    '*********SECURITY**********************************

    str = "SELECT * FROM DISLIB.DISLNKUSR WHERE USR_ID = ? "
    Set USRNAME_QR = rdb.CreateQuery("USRNAME1MICRO", str)
    USRNAME_QR(0) = UID
    Set USRNAME_RS = USRNAME_QR.OpenResultset(1, 1)
    If USRNAME_RS.RowCount > 0 Then
        LOGINUSRNAME = Trim(USRNAME_RS("USERNAME"))
        LOGINTYP = Trim(USRNAME_RS("USRTYPE"))
        LOGINUSRID = Trim(USRNAME_RS("usr_id"))
        cmb_enteryby.Text = (LOGINUSRNAME)
        cmb_enteryby.AddItem LOGINUSRNAME
        'cmb_enteryby.Text = cmb_enteryby.List(0)
        cmb_enteryby.Enabled = False
    Else
        MsgBox "The Security Details of this User is not available!!", vbInformation, "DIS"
        cmd_cancel_Click
        Exit Sub
    End If

    'Set REPWTRIRS = rdb.OpenResultset("SELECT * FROM DISLIB.DISLNKUSR ORDER BY USR_ID", 1, 1)
    str = "SELECT * FROM DISLIB.DISLNKUSR WHERE DEPT_CODE=? ORDER BY USERNAME"
    Set REPWTRN_QR = rdb.CreateQuery("REPWTRN_STRFFF", str)
    If Trim(USRDEPT) = "ALL" Then
        REPWTRN_QR(0) = "GZZ"
    Else
        REPWTRN_QR(0) = Trim(USRDEPT)
    End If
    Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)
    If REPWTRNRS.RowCount > 0 Then
        'cmb_enteryby.clear
        'cmb_addendumby.clear
        cmb_final.clear
        'cmb_provby.clear
        REPWTRNRS.MoveFirst
        While Not REPWTRNRS.EOF
            If (REPWTRNRS("USRTYPE") >= 8 Or REPWTRNRS("USRTYPE") = 6) Then
                'cmb_enteryby.AddItem Trim(REPWTRNRS("USERNAME"))
                cmb_final.AddItem Trim(REPWTRNRS("USERNAME"))
                cmb_addendumby.AddItem Trim(REPWTRNRS("USERNAME"))
                'cmb_provby.AddItem Trim(REPWTRNRS("USERNAME"))
            End If
            REPWTRNRS.MoveNext
        Wend
    Else
        MsgBox "The Security Detials of this User is not available", vbInformation, "DIS"
        cmd_cancel_Click
        Exit Sub
    End If
    '----mrunal
    str = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USR_ID='A" & usr_name & "' "
    Set REPWTRN_QR = rdb.CreateQuery("", str)
   Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)
    If REPWTRNRS.RowCount > 0 Then
    cmb_final.Text = REPWTRNRS("USERNAME")
    End If
    '----
    str = "select * from DISLIB.Dislnkusr where username = ?"
    Set CURUSRDET_QR = rdb.CreateQuery("curusrdetq11MICRO", str)
    CURUSRDET_QR(0) = "ABC"
    Set CURUSRDET_RS = CURUSRDET_QR.OpenResultset(1, 1)
    '
    'STR = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USRTYPE=?"
    'Set SELFINALBY_QR = rdb.CreateQuery("SELFINALBY_STR1MICRO", STR)
    'SELFINALBY_QR(0) = 6
    'Set SELFINALBY_RS = SELFINALBY_QR.OpenResultset(1, 3)
    'If Not SELFINALBY_RS.EOF Then
    '    Do While Not SELFINALBY_RS.EOF
    '       cmb_finalby.AddItem Trim(SELFINALBY_RS("USERNAME"))
    '       SELFINALBY_RS.MoveNext
    '    Loop
    'End If

    '*********SECURITY**********************************

    ''********ICU***************
    str = "SELECT ICDS2 FROM DISLIB.DREQUISITN WHERE REQNO=?"
    Set SELICU_QR = rdb.CreateQuery("SELICU_QR115", str)

    str = "SELECT MEMONO,PAT_CATG FROM RISLIB.RREQUISITN WHERE REQNNO=?" ' AND (PAT_CATG<>'NC' OR PAT_CATG<>'SF')"
    Set SELMEMO_QR = rdb.CreateQuery("SELMEMO_QR115", str)

    ''********ICU***************

    str = "UPDATE RISLIB.RTIMELOG SET COMREPDATE=?,COMREPTIME=?  WHERE REQNNO=?"
    Set UPDRTIMEC_QR = rdb.CreateQuery("QWEQWEQ11111", str)
    ''modified by mrunal--25_2
      str = "UPDATE RISLIB.RTIMELOG SET SAVREPDATE=?,SAVREPTIME=?  WHERE REQNNO=?"
    Set UPDRTIMEC_QR1 = rdb.CreateQuery("", str)
    
      str = "UPDATE RISLIB.RTIMELOG SET MODREPDATE=?,MODREPTIME=? WHERE REQNNO=?"
    Set UPDRTIMEC_QR2 = rdb.CreateQuery("", str)
    'end
    'for addendum
    
    '***** Shraddha 21/03/11 *****
    'SAMCOLDATE  SAMCOLTIME
    str = "UPDATE RISLIB.RTIMELOG SET SAMCOLDATE=?,SAMCOLTIME=? WHERE REQNNO=?"
    Set upd_rtime_qr = rdb.CreateQuery("", str)
    
    str = "SELECT * FROM RISLIB.RTIMELOG WHERE REQNNO=?"
    Set sel_rtimelog_qr = rdb.CreateQuery("", str)
    '*****************************

    str = "UPDATE RISLIB.RTIMELOG SET COMREPDATE=?,COMREPTIME=? WHERE REQNNO=?"
    Set UPDATERTIMELOG = rdb.CreateQuery("UPDATERTIMELOG11", str)
    
    str = "INSERT INTO DISLIB.ADDENDUMMB(REQNO,ADDEND_NO,ADDEND_DT,ADDEND_TM,USER_ID,ADD_MAX_LN) VALUES ( ?,?,?,?,?,?)"
    Set INSDADDENDUM_QR = rdb.CreateQuery("", str)
    
    str = "SELECT  * FROM DISLIB.ADDENDUMMB WHERE REQNO=?"
    Set SELDADDENDUM_QR = rdb.CreateQuery("", str)
    ''''''''
    
    'FOR ANTIBIOTICS
    str = "select antibcode,antibdesc from dislib.dmicbantib WHERE MODALITY=? AND SUBMODALI=? ORDER BY ANTIBDESC"
    Set sel_antiqr = rdb.CreateQuery("micro_select1", str)

    'for dept selection
    str = "select dept_name, dept_code from PATOOLSLIB.ma106"
    Set sel_deptqr = rdb.CreateQuery("", str)

    'FOR MODALITY SELECTION
    str = "select modlitydes,modalitycd from rislib.rmodality WHERE DEPT_CODE=? "
    Set sel_moadlityqr = rdb.CreateQuery("", str)
    
    
    '***** Shraddha 17/02/10 *****
    'Check PRO_CODE in dislib/dmulproce where reqno=?
    'select * from dislib/dmulproce
    str = "select * from dislib.dmulproce where reqno=?"
    Set sel_dreqpc_qr = rdb.CreateQuery("", str)
    '*****************************
    
    '***** Shraddha 22/07/2011 *****
    str = "select * from DISlib.BACOMMENT"
    Set sel_BAcomment_qr = rdb.CreateQuery("", str)
    '*******************************
    
    '************ Shraddha 16/06/2009 ****************
    str = "select modlitydes,modalitycd from rislib.rmodality WHERE DEPT_CODE='GZZ' and  modalitycd=?"
    Set sel_moadlityqr1 = rdb.CreateQuery("", str)
    '**************************************
    'select requisition no.
    'STR = "select * from Rislib.RREQUISITN where caseno=? AND MODALITYCD=? ORDER BY REQNDT DESC"
    'ADDEDD (REQNSTAT=? OR REQNSTAT=?) TO STOP REENTRY OF DATA FOR ALREADY ENTERED RECORD
    'STR = "select * from Rislib.RREQUISITN where caseno=? AND MODALITYCD=? and DEPT_CODE=? AND REQNSTAT='E' ORDER BY REQNDT DESC"
    str = "select * from Rislib.RREQUISITN where caseno=? AND MODALITYCD=? and DEPT_CODE=? AND REQNSTAT='E' ORDER BY REQNDT DESC"
   'str = "select * from Rislib.RREQUISITN where caseno=? AND MODALITYCD=? and DEPT_CODE=? ORDER BY REQNDT DESC"
    Set req_qr = rdb.CreateQuery("micro_select4", str)

    str = "SELECT PATHNO FROM DISLIB.DREQUISITN WHERE REQNO=?"
    Set GZZREQPATH_QR = rdb.CreateQuery("GZZREQPATH_QR", str)

    'select
    'STR = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND MODALITYCD=? AND DEPT_CODE=? AND (REQNSTAT=? OR REQNSTAT=?)ORDER BY REQNDT DESC "
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? and MODALITYCD in ('BA','CM','MO','MB','MC','SE') ORDER BY REQNDT DESC "
    Set SELRREQUISITN_QR = rdb.CreateQuery("micro_SELECT5", str)
'----------Mrunal
' str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND MODALITYCD=? AND  DEPT_CODE=? AND (REQNSTAT=? OR REQNSTAT=?)ORDER BY REQNDT DESC "
'    Set SELRREQUISITN_QR = rdb.CreateQuery("micro_SELECT5", str)
'-------------
    'select other detail for patient
    str = "select * from PATOOLSLIB.ta101 where case_no =?"
    Set SEL_detail_qr = rdb.CreateQuery("micro_select6", str)
    
    '***** Shraddha Pandit 26/11/08
    str = "SELECT * FrOM DISLIB.DMSPONANTI WHERE REQNO=?"
    Set SELSPONANTIBQR = rdb.CreateQuery("SELSPONANTIBQRMB", str)
    
  str = "UPDATE DISLIB.DMSPONANTI SET SPECIMEN=? WHERE REQNO=?"
    Set UPDATERTIMELOG1 = rdb.CreateQuery("", str)
    '**************
    ins_mod_var = "" 'Shraddha 05/06/09
    '*********** Shraddha Pandit 23/01/09 ***********
    str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO = ? "
    Set SEL_DMUL_QR = rdb.CreateQuery("", str)
    '************************************************
    
    'for bed and ward retriving
    str = "select * from PATOOLSLIB.ma104 where case_no =?"
    Set sel_bed_qr = rdb.CreateQuery("micro_select7", str)

    'for referred by dr.
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND REQNNO=? ORDER BY REQNDT DESC"
    Set REFDR_QR = rdb.CreateQuery("micro_select8", str)

    'FOR SELECTING WHICH PROCEDURE HE HAS SELECTED
    str = "SELECT SER_CODE,SER_DET FROM PATOOLSLIB.MA101 WHERE SER_CODE IN(SELECT PRO_CODE FROM DISLIB.DMULPROCE WHERE REQNO= ?)"
    Set sel_proccode_qr = rdb.CreateQuery("micro_select9", str)

    'TO GET MODALITY DESC.
    str = "SELECT MODLITYDES,MODALITYCD FROM RISLIB.RMODALITY WHERE DEPT_CODE=? AND MODALITYCD=?"
    Set SEL_MODALITY_QR = rdb.CreateQuery("micro_SELECT10", str)

    'to get submodality under mycobacteriology
    str = "select DISTINCT(submodali) from dislib.dmicbantib where trim(modality)=?"
    Set sel_submodali_qr = rdb.CreateQuery("micro_select11", str)

    'to get result before modify
    str = "SELECT EPITHECELL,PUSCELL,OFINDING,ZNSTAIN,KOHAMT,GPC,GNC,GPB,GNB,IIPC,ORGSO1,ORGSO2,ORGSO3,REPORT,COMMENTS from  DISLIB.DMOPRISMER where reqno=?"
    Set sel_forupd1_qr = rdb.CreateQuery("micro_SELECT12", str)
 
    'to get antibiotic with values
    str = "SELECT * from  DISLIB.DMBAANTI where reqno=?"
    Set sel_forupd2_qr = rdb.CreateQuery("micro_SELECT13", str)

    str = "select * from DISLIB.DMBAANTI where reqno=? and ANTIBIOTIC=?"
    Set retantibodies_qr = rdb.CreateQuery("retantibiotic", str)

    ' TO INSERT NEW RECORD IF RECORDS ARE NOT FOUND DURING MODIFY
    str = "SELECT * from  DISLIB.DMBAANTI where reqno=? AND antibiotic=?"
    Set sel_formod2_qr = rdb.CreateQuery("micro_SELECTmod13", str)


    'TO GET FOR VIEW OF URINE RESULT
    str = "SELECT * from  DISLIB.DMURE where reqno=?"
    Set sel_forupd3_qr = rdb.CreateQuery("micro_SELECT14", str)

    'TO GET FOR VIEW OF URINE RESULT
    str = "SELECT * from  DISLIB.DMSRE where reqno=?"
    Set sel_forupd4_qr = rdb.CreateQuery("micro_SELECT15", str)

    'TO GET FOR VIEW OF other RESULT
    str = "SELECT * from  DISLIB.dmbsemcmo where reqno=?"
    Set sel_forupd5_qr = rdb.CreateQuery("micro_SELECT16", str)

    'to get result of myco
    str = "SELECT * from  DISLIB.dmycobacte where reqno=?"
    Set sel_forupd6_qr = rdb.CreateQuery("micro_SELECT17", str)

    'to get micro no which is path no. and previous micro no. in report
    str = "SELECT * from  DISLIB.Drequisitn where reqno=?"
    Set sel_micro_qr = rdb.CreateQuery("micro_SELECT18", str)

    'to check status of commitflag
    str = "select * from rislib.rrequisitn where reqnno=?"
    Set sel_ris_qr = rdb.CreateQuery("micro_select19", str)

    'to get entered by name
    str = "select * from dislib.dissec where enterby='YES'AND MODALITY='BA'"
    Set sel_ENTER_qr = rdb.CreateQuery("seldissecentr_qr0091", str)

    'to ge fianl by list name
    str = "select * from dislib.dissec where FINALby='YES' AND MODALITY='BA'"
    Set sel_FINAL_qr = rdb.CreateQuery("seldissecFIN_qr0091", str)

    'for staff query
    str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=? AND EMPCODE=?"
    Set SEL_dreq_staff_qr = rdb.CreateQuery("micro_select", str)

    'STR = "select * from dislib.dissec where PROVby='YES' AND MODALITY='BA'"
    'Set seldissecprov_qr = rdb.CreateQuery("seldissecPROV_qr1", STR)

    'STR = "select * from dislib.dissec where ADDENDUMby='YES'AND MODALITY='BA'"
    'Set seldissecadd_qr = rdb.CreateQuery("seldissecADD_qr1", STR)

    'STR = "select * from dislib.dissec where GROSSby='YES'AND MODALITY='BA'"
    'Set seldissecgross_qr = rdb.CreateQuery("seldissecGROSS_qr1", STR)

    '###################################    INSERT QUERY    ##############################
    'insert into table for modality except 1 and 5 record
    'OVER
    str = "insert into DISLIB.dmbsemcmo(reqno,refserv,testperfor,result)values(?,?,?,?)"
    Set ins_0_qr = rdb.CreateQuery("micro_insert0", str)
    'insert record for madality 1
    'over
    str = "insert into DISLIB.dmoprismer(reqno,epithecell,puscell,ofinding,znstain,kohamt,gpc,gnc,gpb,gnb,iipc,orgso1,orgso2,orgso3,REPORT,comments) values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
    Set ins_1_qr = rdb.CreateQuery("micro_insert1", str)
    'for inserting effect of antibiotics for modality 1
    'over
    str = "insert into DISLIB.dmbaanti(reqno,antibiotic,antibio1,antibio2,antibio3) values(?,?,?,?,?)"
    Set ins_2_qr = rdb.CreateQuery("micro_insert2", str)

    str = "insert into dislib.GZZSEROLOG values(?,?,?,?,?,?,?,?,?,?,?)"
    Set insGZZqr = rdb.CreateQuery("insertIntoGZZ", str)
    'insert urine result
    'over
    'str = "insert into DISLIB.dmure (reqno,color,sp,ph,protein,glucose,ketones,bulirubin,urobilinog,blood,micexam,wbc,rbc,sqepcell,bacteria,casts,others,comments)values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
    'Commented and added by Shraddha on 21/03/11 *****
    str = "insert into DISLIB.dmure (reqno,color,sp,ph,protein,glucose,ketones,bulirubin,urobilinog,blood,micexam,wbc,rbc,sqepcell,bacteria,casts,others,comments,NITRITE)values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
    Set ins_4_qr = rdb.CreateQuery("micro_insert4", str)
    'insert stool result
    'over
    str = "insert into DISLIB.dmsre (reqno,color,consistenc,mucus,blood,parasties,occuttblod,ova,cysts,larvae,puscell,rbc,macophages,starch,fat,others,comments) values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
    Set ins_5_qr = rdb.CreateQuery("micro_insert5", str)

    'insert into mycobacteriology table
    str = "insert into DISLIB.dmycobacte (reqno,stain,organism,report,comment) values(?,?,?,?,?)"
    Set ins_6_qr = rdb.CreateQuery("micro_insert6", str)

    ''INSERT  DUMMY REQNO INTO DMURE TABLE
    'STR = "insert into DISLIB.dmure (reqno) values(?)"
    'Set ins_7_qr = rdb.CreateQuery("micro_insert7", STR)
    '
    ''insert DUMMY REQNO IN stool TABLE
    'STR = "insert into DISLIB.dmsre (reqno) values(?)"
    'Set ins_8_qr = rdb.CreateQuery("micro_insert8", STR)

    '###################################    UPDATE QUERY    ##############################
    'update queries
    'update reqstat if he/she successfully entered data
    str = "update rislib.rrequisitn set reqnstat='T' where caseno=? and reqnno=?"
    Set upd_11_qr = rdb.CreateQuery("micro_update0", str)

    str = "update dislib.dmbsemcmo set refserv=? ,testperfor=?,result=? where reqno=?"
    Set upd_0_qr = rdb.CreateQuery("micro_update1", str)

    
    'UPDATE record for madality 1
    str = "update dislib.dmoprismer set epithecell=? ,puscell=?,ofinding=?,znstain=?,kohamt=?,gpc=?,gnc=?,gpb=?,gnb=?,iipc=?,orgso1=?,orgso2=?,orgso3=?,REPORT=?,comments=? where reqno=? "
    Set upd_1_qr = rdb.CreateQuery("micro_update2", str)

    'for UPDATE effect of antibiotics for modality 1
    str = "update dislib.dmbaanti set antibio1=?,antibio2=?,antibio3=? where reqno=? and antibiotic=? "
    Set upd_2_qr = rdb.CreateQuery("micro_update3", str)

    str = "update dislib.GZZSEROLOG set TESTPERF=?,Result=?,PROCALC=?,PROCAL_CN=?,TESTRANGE=?,INTERPRETN=?,METHOD=?,LIMITATION=?,FLDBLD=?,FLD1=? where REQNO=?"
    Set UpdGZZqr = rdb.CreateQuery("updateGZZ", str)
    
    'for UPDATE effect of antibiotics for modality 1

    str = "update dislib.dmbaanti set antibio1=?,antibio2=?,antibio3=?,antibiotic=?  where reqno=? "
    Set upd_12_qr = rdb.CreateQuery("micro_update13", str)

    'update urine result
    'str = "update dislib.dmure set color=?,sp=?,ph=?,protein=?,glucose=?,ketones=?,bulirubin=?,urobilinog=?,blood=?,micexam=?,wbc=?,rbc=?,sqepcell=?,bacteria=?,casts=?,others=?,comments=? where reqno=? "
    'Commented and added by Shraddha on 21/03/11
    str = "update dislib.dmure set color=?,sp=?,ph=?,protein=?,glucose=?,ketones=?,bulirubin=?,urobilinog=?,blood=?,micexam=?,wbc=?,rbc=?,sqepcell=?,bacteria=?,casts=?,others=?,comments=?, NITRITE =? where reqno=? "
    Set upd_4_qr = rdb.CreateQuery("micro_update4", str)

    'update stool result
    str = "update dislib.dmsre set color=?,consistenc=?,mucus=?,blood=?,parasties=?,occuttblod=?,ova=?,cysts=?,larvae=?,puscell=?,rbc=?,macophages=?,starch=?,fat=?,others=?,comments=?  where reqno=? "
    Set upd_5_qr = rdb.CreateQuery("micro_update5", str)

    'update commit flag
    str = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
    Set upd_6_qr = rdb.CreateQuery("micro_update6", str)

    'update mycobacteriology table
    str = "update DISLIB.dmycobacte set stain=? ,organism=?,report=? ,comment=? where reqno=?"
    Set upd_7_qr = rdb.CreateQuery("micro_update7", str)

    'to update rislib for enterby and enter by
    str = "update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?,freportby=?,freportdt=? WHERE REQNNO=?"
    Set upd_8_qr = rdb.CreateQuery("micro_update_8", str)

    '''Done By KANCHAN

    str = "delete from dislib.dmbaanti where reqno=?"
    Set del_dmbaanti_qr = rdb.CreateQuery("micro_delete1", str)
    'delete for specific antibiotic
    str = "delete from dislib.dmbaanti where reqno=? and antibiotic=?"
    Set del_dmbaanti_qr1 = rdb.CreateQuery("micro_delete2", str)

    str = "SELECT * FROM DISLIB.DMYCOBACTE WHERE REQNO=?"
    Set CHECKREC_QR = rdb.CreateQuery("CHECKRECORD", str)

    str = "SELECT REPORT FROM DISLIB.DMYCOBACTE WHERE REQNO=?"
    Set REPORT_QR = rdb.CreateQuery("REPORT", str)

    ''DECOMMIT
    str = "UPDATE RISLIB.RREQUISITN SET REQNSTAT=?,COMMITFLAG=? WHERE REQNNO=?"
    Set UPDRREQ_QR = rdb.CreateQuery("UPDATERREQ", str)

    str = "SELECT * FROM DISLIB.DMBSEMCmo WHERE REQNO=?"
    Set SELDMBSEMCOM_QR = rdb.CreateQuery("SELDMBSEMCOM111", str)

    'grid_anti.AddItem
End If
btnenabl ("INI")
new_flag = False
VIEW_FLAG = False
mod_flag = False
GRIDSCROLL = False

''**********
'25/10/2011
Dim MICROMO_QR As rdoQuery
Dim MICROMO_RS As rdoResultset

str = "SELECT * FROM DISLIB.MICROMO"
Set MICROMO_QR = rdb.CreateQuery("", str)

Set MICROMO_RS = MICROMO_QR.OpenResultset(1, 3)

While Not MICROMO_RS.EOF
    cmb_motemplet.AddItem Trim(MICROMO_RS("TESTNAME"))
    MICROMO_RS.MoveNext
Wend

'************


''**********
'10/03/2014
Dim MICRdrug_QR As rdoQuery
Dim MICRdrug_RS As rdoResultset

str = "SELECT * FROM DISLIB.MICR_DRUG2 "
Set MICRdrug_QR = rdb.CreateQuery("", str)

Set MICRdrug_RS = MICRdrug_QR.OpenResultset(1, 3)

While Not MICRdrug_RS.EOF
    cmbmicrodrug.AddItem Trim(MICRdrug_RS("DESCRIP"))
    MICRdrug_RS.MoveNext
Wend

'************

'******************* Shraddha Pandit 02/07/09 *****************

Dim org_QR As rdoQuery
Dim org_RS As rdoResultset
Dim sql As String

sql = "SELECT * FROM DISLIB.organism order by ORGANISM"
Set org_QR = rdb.CreateQuery("", sql)
Set org_RS = org_QR.OpenResultset(1, 3)

If Not org_RS.EOF Then
      
    While Not org_RS.EOF
        
'        cmb_org1.AddItem org_RS("ORGANISM")
'        cmb_org2.AddItem org_RS("ORGANISM")
'        cmb_org3.AddItem org_RS("ORGANISM")
    '***** 02/07/09
        cmb_org1.AddItem Trim(org_RS("ORGANISM"))
        cmb_org1.AddItem Trim(org_RS("ORGANISM")) & "$" & "Heavy"
        cmb_org1.AddItem Trim(org_RS("ORGANISM")) & "$" & "Poor"
        cmb_org1.AddItem Trim(org_RS("ORGANISM")) & "$" & "Moderate"
        cmb_org1.AddItem Trim(org_RS("ORGANISM")) & "$" & "Significant"
        
        cmb_org2.AddItem Trim(org_RS("ORGANISM"))
        cmb_org2.AddItem Trim(org_RS("ORGANISM")) & "$" & "Heavy"
        cmb_org2.AddItem Trim(org_RS("ORGANISM")) & "$" & "Poor"
        cmb_org2.AddItem Trim(org_RS("ORGANISM")) & "$" & "Moderate"
        cmb_org2.AddItem Trim(org_RS("ORGANISM")) & "$" & "Significant"
        
        cmb_org3.AddItem Trim(org_RS("ORGANISM"))
        cmb_org3.AddItem Trim(org_RS("ORGANISM")) & "$" & "Heavy"
        cmb_org3.AddItem Trim(org_RS("ORGANISM")) & "$" & "Poor"
        cmb_org3.AddItem Trim(org_RS("ORGANISM")) & "$" & "Moderate"
        cmb_org3.AddItem Trim(org_RS("ORGANISM")) & "$" & "Significant"
    '***** 02/07/09
        org_RS.MoveNext
        
    Wend
        
End If
org_RS.Close
'**************************************************************

'***** Shraddha 22/07/2011 *****
'Set sel_BAcomment_rs = sel_BAcomment_qr.OpenResultset(1, 3)
'While Not sel_BAcomment_rs.EOF
'    txt_comment.AddItem Trim(sel_BAcomment_rs("COMMENT"))
'    sel_BAcomment_rs.MoveNext
'Wend
'*******************************


'set_position 'to set the position of text_field on grid.
End Sub

Private Sub set_visible(frame_name As String)
If frame_name = "All" Then
    frm_bacteria.Visible = False
    frm_others.Visible = False
    frm_clinic.Visible = False
    frm_myco.Visible = False
    Frame3.Visible = False 'Shraddha 11/11/08
    frmSerology.Visible = False
ElseIf frame_name = "frm_bacteria" Then
    frm_bacteria.Visible = True
    Frame3.Visible = False
    frm_others.Visible = False
    frm_clinic.Visible = False
    frm_myco.Visible = False
    frmSerology.Visible = False
ElseIf frame_name = "frm_others" Then
    frm_others.Visible = True
    frm_bacteria.Visible = False
    frm_clinic.Visible = False
    frm_myco.Visible = False
    frmSerology.Visible = False
    Frame3.Visible = False
ElseIf frame_name = "frm_myco" Then
    frm_others.Visible = False
    frm_bacteria.Visible = False
    frm_clinic.Visible = False
    frm_myco.Visible = True
    
    If new_flag = True Then
        cmb_antiboitic_mb.Visible = False
        txt_c_1.Visible = False
    End If
    frmSerology.Visible = False
    Frame3.Visible = False
ElseIf frame_name = "frmSerology" Then
    frmSerology.Visible = True
    frm_others.Visible = False
    frm_bacteria.Visible = False
    frm_clinic.Visible = False
    frm_myco.Visible = False
    Frame3.Visible = False
    sel_dreqpc_qr(0) = Trim(mod_reqno)
Set sel_dreqpc_rs = sel_dreqpc_qr.OpenResultset(1, 3)
If sel_dreqpc_rs("PRO_CODE") = "G262" Or sel_dreqpc_rs("PRO_CODE") = "G204" Then
 Frame3.Visible = True
End If
ElseIf frame_name = "frame3" Then
    frmSerology.Visible = False
    frm_others.Visible = False
    frm_bacteria.Visible = False
    frm_clinic.Visible = False
    frm_myco.Visible = False
    Frame3.Visible = True
    
Else
    frm_clinic.Visible = True
    frm_bacteria.Visible = False
    frm_others.Visible = False
    frm_myco.Visible = False
    frmSerology.Visible = False
    Frame3.Visible = False
End If
End Sub







'Private Sub grid_anti_Click()
'If grid_anti.Col = 1 Or grid_anti.Col = 2 Or grid_anti.Col = 3 Then
'    'If new_flag Then
'       If Len(txt_organ1.Text) > 1 Then
'          If Len(TXT_ORGAN2.Text) > 1 Then
'               If Len(TXT_ORGAN3.Text) > 1 Then
'                    'three organisms
'                    txt_1 = CStr(grid_anti.TextMatrix(grid_anti.row, 1))
'                    txt_2 = CStr(grid_anti.TextMatrix(grid_anti.row, 2))
'                    txt_3 = CStr(grid_anti.TextMatrix(grid_anti.row, 3))
'                    grid_anti.Col = 1
'                    txt_1.Top = grid_anti.CellTop + grid_anti.Top
'                    grid_anti.Col = 2
'                    txt_2.Top = grid_anti.CellTop + grid_anti.Top
'                    grid_anti.Col = 3
'                    txt_3.Top = grid_anti.CellTop + grid_anti.Top
''                    If grid_anti.Col = 1 Then
''                        txt_1.Visible = True
''                        txt_2.Visible = False
''                        txt_3.Visible = False
''                        txt_1.SetFocus
''                    ElseIf grid_anti.Col = 2 Then
''                        txt_1.Visible = False
''                        txt_2.Visible = True
''                        txt_3.Visible = False
''                        txt_2.SetFocus
''                    ElseIf grid_anti.Col = 3 Then
''                        txt_1.Visible = False
''                        txt_2.Visible = False
''                        txt_3.Visible = True
''                        txt_3.SetFocus
''                    End If
'               Else
'                    'two organisms
'                     txt_1.Text = CStr(grid_anti.TextMatrix(grid_anti.row, 1))
'                     txt_2.Text = CStr(grid_anti.TextMatrix(grid_anti.row, 2))
'                     grid_anti.Col = 1
'                     'txt_1.Visible = False
'                     txt_1.Top = grid_anti.CellTop + grid_anti.Top
'                     grid_anti.Col = 2
'                     txt_2.Top = grid_anti.CellTop + grid_anti.Top
''                     If grid_anti.Col = 1 Then
''                        txt_1.Visible = True
''                        txt_2.Visible = False
''                        txt_1.SetFocus
''                    ElseIf grid_anti.Col = 2 Then
''                        txt_1.Visible = False
''                        txt_2.Visible = True
''                        txt_2.SetFocus
''                    End If
'                     txt_3.Visible = False
'               End If
'          Else
'            'only one organisms
'            txt_1.Text = CStr(grid_anti.TextMatrix(grid_anti.row, 1))
'            grid_anti.Col = 1
'            txt_1.Top = grid_anti.CellTop + grid_anti.Top
'
''            If grid_anti.Col = 1 Then
''                        txt_1.Visible = True
''                        txt_1.SetFocus
''            End If
''                        txt_2.Visible = False
''                        txt_3.Visible = False
'          End If
'       End If
'   End If
''End If
'End Sub

Private Sub grid_c_anti_Click()
'    If grid_c_anti.col = 1 Or grid_c_anti.col = 2 Then
'        txt_c_1.Text = CStr(grid_c_anti.TextMatrix(grid_c_anti.row, 1))
'        grid_c_anti.col = 1
'        txt_c_1.Top = grid_c_anti.CellTop + grid_c_anti.Top
'        If txt_c_modflg = True Then
'            txt_c_1.Visible = True
'        End If
'    End If
    'Commented and added
    '***** Shraddha 11/05/11 *****
    GRIDSCROLL = False
    If MODE = "MODIFY" Or MODE = "NEW" Then
     '   If (Len(Trim(txt_c_1.Text)) > 0) Then
            If grid_c_anti.row = 0 Then
                MsgBox ("Please select proper location....!"), vbOKOnly
            Else
                If MODE = "MODIFY" Or MODE = "NEW" Then
                    If grid_c_anti.col = 0 Then
                        cmb_antiboitic_mb.Visible = True
                        cmb_antiboitic_mb.Enabled = True
                        cmb_antiboitic_mb.SetFocus
                        txt_c_1.Visible = False '12/05/11
                    ElseIf grid_c_anti.col = 1 Then
                        If Len(Trim(grid_c_anti.TextMatrix(grid_dmbanti.row, 0))) > 0 Then
                            txt_c_1.Visible = True
                            txt_c_1.Enabled = True
                            txt_c_1.SetFocus
                            
                            '***** Shraddha ******

                            grid_c_anti.col = 1
                            txt_c_1.Move grid_c_anti.Left + grid_c_anti.CellLeft, grid_c_anti.Top + grid_c_anti.CellTop, grid_c_anti.CellWidth
                            txt_c_1.Text = grid_c_anti.TextMatrix(grid_c_anti.row, 1)
                            txt_c_1.Visible = True
                            txt_c_1.SetFocus
                            '**********************
                        Else
                            MsgBox ("Please select antibiotic first....!"), vbOKOnly
                            cmb_antiboitic_mb.Visible = True
                            cmb_antiboitic_mb.Enabled = True
                            cmb_antiboitic_mb.SetFocus
                            Call GridControls_mb("VISIBLE", 0)
                            Call movecontrols_MB(, 0)
                            Exit Sub
                        End If
                    End If
                    
                    Call GridControls_mb("VISIBLE", grid_dmbanti.col)
                    Call movecontrols_MB(, grid_dmbanti.col)
                    
                Else
                    GridControls_mb ("INVISIBLE")
                End If
            End If
        'Else
         '   MsgBox ("Atleast 1 organism is needed...!"), vbOKOnly
         '   cmb_organisms.SetFocus
 '       End If
    End If
End Sub

Private Sub grid_dmbanti_Click()

''''' Done by Kanchan

GRIDSCROLL = False
If MODE = "MODIFY" Or MODE = "NEW" Then

'********* Shraddha 30/06/09 ****************
'If (Len(Trim(txt_organ1.Text)) > 0 Or Len(Trim(TXT_ORGAN2.Text)) > 0 Or Len(Trim(TXT_ORGAN3.Text)) > 0) Then
If (Len(Trim(cmb_org1.Text)) > 0 Or Len(Trim(cmb_org2.Text)) > 0 Or Len(Trim(cmb_org3.Text)) > 0) Then
'********************************************
If grid_dmbanti.row = 0 Then
    MsgBox ("Please select proper location....!"), vbOKOnly
Else
    If MODE = "MODIFY" Or MODE = "NEW" Then
        If grid_dmbanti.col = 0 Then
            cmb_antibiotic.Visible = True
            cmb_antibiotic.Enabled = True
            cmb_antibiotic.SetFocus
        ElseIf grid_dmbanti.col = 1 Then
            If Len(Trim(grid_dmbanti.TextMatrix(grid_dmbanti.row, 0))) > 0 Then
                txt_1.Visible = True
                txt_1.Enabled = True
                txt_1.SetFocus
            Else
                MsgBox ("Please select antibiotic first....!"), vbOKOnly
                cmb_antibiotic.Visible = True
                cmb_antibiotic.Enabled = True
                cmb_antibiotic.SetFocus
                Call GridControls("VISIBLE", 0)
                Call movecontrols(, 0)
                Exit Sub
            End If
        ElseIf grid_dmbanti.col = 2 Then
            txt_2.Visible = True
            txt_2.Enabled = True
            txt_2.SetFocus
        ElseIf grid_dmbanti.col = 3 Then
            txt_3.Visible = True
            txt_3.Enabled = True
            txt_3.SetFocus
         End If
              
            Call GridControls("VISIBLE", grid_dmbanti.col)
            Call movecontrols(, grid_dmbanti.col)
    Else
            GridControls ("INVISIBLE")
                  
    End If
    
 End If
 
Else
    MsgBox ("Atleast 1 organism is needed...!"), vbOKOnly
    '***** Commented and added by Shraddha on 30/06/09 ************
    'txt_organ1.SetFocus
    cmb_org1.SetFocus
    '*************************************************************
End If ''''CHECK FOR ORGANISM

End If   ''''MODE
End Sub

Private Sub grid_dmbanti_Scroll()
    If GRIDSCROLL = False Then
        GridControls ("INVISIBLE")
        grid_dmbanti.SetFocus
    End If
End Sub

Private Sub MaskEdBox1_LostFocus()
'On Error GoTo MSG:
If Trim(MaskEdBox1.Text) = "" Then 'Shraddha Pandit 17/11/08
    'btnenabl "NEW"
    Exit Sub
End If

If Opt_patient.Value Then
    MaskEdBox1.Text = UCase(padzero(5, "C", MaskEdBox1.Text, 3))
    If new_flag Then
            cmb_reqno.clear
            req_qr(0) = MaskEdBox1.FormattedText
            req_qr(1) = Trim(txt_modcode.Text)
            req_qr(2) = Trim(txt_deptcode.Text)
            Set req_rs = req_qr.OpenResultset(1, 3)
        If req_rs.RowCount > 0 Then
           ' no_req = True
            
            'ADDED BY MILIND ON 16-10-06 FOR DISPLAYING REQNO WITH LABNO
            While Not req_rs.EOF
                GZZREQPATH_QR(0) = req_rs("REQNNO")
                Set GZZREQPATH_RS = GZZREQPATH_QR.OpenResultset(1, 3)
                While Not GZZREQPATH_RS.EOF
                    cmb_reqno.AddItem req_chglen(req_rs("reqNno") + "$" + GZZREQPATH_RS("PATHNO"))
                    GZZREQPATH_RS.MoveNext
                Wend
                req_rs.MoveNext
            Wend
            req_rs.Close
            GZZREQPATH_RS.Close
            'UPTO HERE
        Else
            no_req = False
            req_rs.Close
            '131008 GZZREQPATH_RS.Close
            cmb_reqno.clear
            MsgBox "No Requisition for this Modality", vbOKOnly
            btnenabl "cancel"
            set_flag "asd"
            set_visible "All"
            '131008 GZZREQPATH_RS.Close
            Exit Sub
            
        End If 'end of rowcount
        ' cmb_reqno.SetFocus
        'SendKeys "^{f4}"
    End If  'end of new_flag
    
        If VIEW_FLAG Then
            new_flag = False
            SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
'            SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
'            SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
'            SELRREQUISITN_QR(1) = "T"
'            SELRREQUISITN_QR(2) = "F"
            Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
            If SELRREQUISITN_RS.RowCount > 0 Then
               ' no_req = True
                cmb_reqno.clear
'                GZZREQPATH_QR(0) = SELRREQUISITN_RS("REQNNO")
'                Set GZZREQPATH_RS = GZZREQPATH_QR.OpenResultset(1, 3)
                'ADDED BY MILIND ON 16-10-06 FOR DISPLAYING REQNO WITH LABNO
                While Not SELRREQUISITN_RS.EOF
                    GZZREQPATH_QR(0) = SELRREQUISITN_RS("REQNNO")
                    Set GZZREQPATH_RS = GZZREQPATH_QR.OpenResultset(1, 3)
                    While Not GZZREQPATH_RS.EOF
                        cmb_reqno.AddItem req_chglen(SELRREQUISITN_RS("REQNNO")) + "$" + GZZREQPATH_RS("pathno")
                             GZZREQPATH_RS.MoveNext
                    Wend
                        SELRREQUISITN_RS.MoveNext
                Wend
                SELRREQUISITN_RS.Close
                GZZREQPATH_RS.Close
                'UPTO HERE
            Else
                no_req = False
                SELRREQUISITN_RS.Close
                'GZZREQPATH_RS.Close 'Shraddha Pandit 18/11/08
                MsgBox "No Requisition for this Modality", vbOKOnly
                btnenabl "cancel"
                set_flag "asd"
                set_visible "All"
                
                Exit Sub
            End If
        End If
ElseIf opt_staff.Value = True Then
    If new_flag Then
            cmb_reqno.clear
                req_qr(0) = MaskEdBox1.FormattedText
                req_qr(1) = Trim(txt_modcode.Text)
                req_qr(2) = Trim(txt_deptcode.Text)
                Set req_rs = req_qr.OpenResultset(1, 3)
            If req_rs.RowCount > 0 Then
                'no_req = True
                
                'ADDED BY MILIND ON 16-10-06 FOR DISPLAYING REQNO WITH LABNO
                While Not req_rs.EOF
                    GZZREQPATH_QR(0) = req_rs("REQNNO")
                    Set GZZREQPATH_RS = GZZREQPATH_QR.OpenResultset(1, 3)
                    While Not GZZREQPATH_RS.EOF
                        cmb_reqno.AddItem req_chglen(req_rs("reqNno")) + "$" + GZZREQPATH_RS("pathno")
                        GZZREQPATH_RS.MoveNext
                    Wend
                        req_rs.MoveNext
                Wend
                'UPTO HERE
            Else
                cmb_reqno.clear
                no_req = False
                MsgBox "No Requisition for this Modality", vbOKOnly
                btnenabl ("INI")
                Exit Sub
            End If
    End If
    If VIEW_FLAG Then
                cmb_reqno.clear
                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
            'End If
            If SELRREQUISITN_RS.RowCount > 0 Then
                'no_req = True
                'ADDED BY MILIND ON 16-10-06 FOR DISPLAYING REQNO WITH LABNO
                While Not SELRREQUISITN_RS.EOF
                        GZZREQPATH_QR(0) = SELRREQUISITN_RS("REQNNO")
                        Set GZZREQPATH_RS = GZZREQPATH_QR.OpenResultset(1, 3)
                    While Not GZZREQPATH_RS.EOF
                        cmb_reqno.AddItem req_chglen(SELRREQUISITN_RS("REQNNO")) + "$" + GZZREQPATH_RS("pathno")
                        GZZREQPATH_RS.MoveNext
                    Wend
                        SELRREQUISITN_RS.MoveNext
                Wend
                SELRREQUISITN_RS.Close
                GZZREQPATH_RS.Close
                'UPTO HERE
             Else
                no_req = False
                MsgBox " No Requisition for this Modality"
                SELRREQUISITN_RS.Close
                GZZREQPATH_RS.Close
                btnenabl ("INI")
                Exit Sub
             End If
        End If
End If
 Exit Sub
MSG:
    MsgBox ERROR(Err)
End Sub


Private Sub MSFlexGrid1_Click()
'*********** Shraddha Pandit 27/10/08 ***********
If MSFlexGrid1.col = 1 Then 'Shraddha 26/02/10 Hospital's 70th Annual day
    text1.Visible = True
    text1.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
    text1.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
    text1.Visible = True
    text1.Text = ""
    text1.SetFocus
    cmbmicrodrug.Visible = False
ElseIf MSFlexGrid1.col = 3 Then '10/03/2014
    cmbmicrodrug.Visible = True
    cmbmicrodrug.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
    cmbmicrodrug.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
    cmbmicrodrug.Visible = True
    cmbmicrodrug.Text = "Select"
    cmbmicrodrug.SetFocus
     text1.Visible = False
End If
''MSFlexGrid1.col = 3
'***********************************************
    
End Sub

Private Sub Opt_patient_Click()
    If Opt_patient.Value = True Then
       lbl_caseno.Caption = "Case No."
       MaskEdBox1.Mask = "??>/#####"
       'MaskEdBox1.SetFocus
    End If
    
End Sub

Private Sub opt_staff_Click()
If opt_staff.Value = True Then
    lbl_caseno.Caption = "CC No."
    MaskEdBox1.Mask = "#######"
    'MaskEdBox1.SetFocus
End If
End Sub

Private Sub Text1_Change()
   If MSFlexGrid1.col = 1 Then
        If text1.Text <> "" Then
           '' MSFlexGrid1.CellBackColor = vbGreen
        End If
        MSFlexGrid1.Text = text1.Text
    End If
  
End Sub
''10/03/2014
Private Sub cmbmicrodrug_Change()
If MSFlexGrid1.col = 3 Then
        If cmbmicrodrug.Text <> "" Then
            ''MSFlexGrid1.CellBackColor = vbGreen
        End If
        MSFlexGrid1.Text = cmbmicrodrug.Text
End If


End Sub

Private Sub Text1_GotFocus()
 If MSFlexGrid1.col = 3 Then
'            Combo1.Top = text1.CellTop + text1.Top
'            Combo1.Left = text1.CellLeft + text1.Left
'            Combo1.Text = text1.TextMatrix(text1.row, text1.col)
 End If
End Sub
''10/03/2014
Private Sub cmbmicrodrug_KeyPress(KeyAscii As Integer)

If MSFlexGrid1.row <= 10 Then
    If MSFlexGrid1.col = 3 Then
        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) = cmbmicrodrug.Text
        If KeyAscii = 13 And MSFlexGrid1.row = 10 Then
            cmbmicrodrug.Visible = False
            'cmd_commit.SetFocus
        End If
        If KeyAscii = 13 And MSFlexGrid1.row < 10 Then
       ' text1_LostFocus
        'get
                text1.Visible = True
                MSFlexGrid1.row = MSFlexGrid1.row + 1
                MSFlexGrid1.col = 1
                text1.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                text1.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
                text1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, MSFlexGrid1.col)
                cmbmicrodrug.Visible = False
                text1.SetFocus
        End If
        
    End If
    'cmd_commit.SetFocus
End If
End Sub
'***** Shraddha 02/03/10 *****
Private Sub text1_KeyPress(KeyAscii As Integer)
'If MSFlexGrid1.row <> 8 Then
'    If MSFlexGrid1.col = 1 Then
'        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1) = Text1.Text
'        If KeyAscii = 13 And MSFlexGrid1.row < 17 Then
'            MSFlexGrid1.row = MSFlexGrid1.row + 1
'            Text1.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
'            Text1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, MSFlexGrid1.col)
'        End If
'    ElseIf MSFlexGrid1.col = 3 Then
'        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) = Text1.Text
'        If KeyAscii = 13 And MSFlexGrid1.row < 17 Then
'            MSFlexGrid1.row = MSFlexGrid1.row + 1
'            Text1.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
'            Text1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, MSFlexGrid1.col)
'        End If
'    End If
'End If

If MSFlexGrid1.row <= 10 Then
    If MSFlexGrid1.col = 1 Then
        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1) = text1.Text
        If KeyAscii = 13 And MSFlexGrid1.row <= 10 Then
            'MSFlexGrid1.row = MSFlexGrid1.row + 1
            'MSFlexGrid1.col = MSFlexGrid1.col + 2
             'GetResult
            text1.Visible = False
            cmbmicrodrug.Visible = True
            
            MSFlexGrid1.col = 3
            cmbmicrodrug.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            cmbmicrodrug.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
            cmbmicrodrug.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, MSFlexGrid1.col)
            cmbmicrodrug.SetFocus
        End If
'    ElseIf MSFlexGrid1.col = 3 Then
'        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) = text1.Text
'        If KeyAscii = 13 And MSFlexGrid1.row = 8 Then
'            text1.Visible = False
'            'cmd_commit.SetFocus
'        End If
'        If KeyAscii = 13 And MSFlexGrid1.row < 8 Then
'       ' text1_LostFocus
'        'get
'
'                MSFlexGrid1.row = MSFlexGrid1.row + 1
'                MSFlexGrid1.col = 1
'                text1.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
'                text1.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
'                text1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, MSFlexGrid1.col)
'
'        End If
'
  End If
    'cmd_commit.SetFocus
End If
End Sub
'**********************************************************


Private Sub Timer1_Timer()

'datetime
'txt_time.Text = time1
'VAR_REQNNO = cmb_reqno.Text
'cmb_reqno.Text = VAR_REQNNO
End Sub
Function GetResult()
'addde by shweta 05-03-2014
    Dim micr_qr1 As rdoQuery
    Dim micr_rs1 As rdoResultset
     
    Dim micr_qr2 As rdoQuery
    Dim micr_rs2 As rdoResultset
    Dim S As String, r As String, d As String, I As String
    Dim S1 As String, R1 As String, D1 As String, I1 As String
    Dim S2 As String, R2 As String, D2 As String, I2 As String
    Dim str As String, FLD1 As String, SPC1 As String
'    str = "SELECT * FROM DISLIB.MICR_DRUG2 ORDER BY FLD_NAME"
'    Set micr_qr2 = rdb.CreateQuery("", str)
'    Set micr_rs2 = micr_qr2.OpenResultset(1, 1)
'    If micr_rs2.RowCount > 0 Then
'    While micr_rs2.EOF = False
'    FLD1 = Trim(micr_rs2("FLD_NAME"))
'    SPC1 = Trim(micr_rs2("SPCHAR"))
    
        str = "SELECT * FROM DISLIB.MICR_DRUG1 WHERE DRUG_NAME = ?"
        Set micr_qr1 = rdb.CreateQuery("", str)
        micr_qr1(0) = MSFlexGrid1.TextMatrix(1, 0)
        Set micr_rs1 = micr_qr1.OpenResultset(1, 1)
        If micr_rs1.RowCount > 0 Then
        S = Trim(micr_rs1("SUSCE"))
        Dim kp
            For kp = 0 To Len(S) - 1
                If IsNumeric(Mid(S, kp + 1, 1)) = True Then
                Exit For
                Else
                S1 = Mid(S, kp + 1, 1)
                S2 = S2 & S1
                End If
            Next
            If S2 = "<" Then
            ElseIf S2 = ">" Then
            ElseIf S2 = "<=" Then
            
            ElseIf S2 = "=>" Then
            
            End If
            If S < 0.2 Then
            MsgBox "yes"
            Else
            MsgBox "no"
            End If
            
        d = Trim(micr_rs2("DOSE"))
        I = Trim(micr_rs2("INTERMED"))
        r = Trim(micr_rs2("RESISTANT"))
            
        
        End If
'    micr_rs2.MoveNext
'    Wend
      
'    End If
    
    '''till here
End Function
Private Sub txt_1_Change()
If Len(Trim(txt_1.Text)) = 1 And Not (UCase(Trim(txt_1.Text)) = "S" Or UCase(Trim(txt_1.Text)) = "I" Or UCase(Trim(txt_1.Text)) = "R") Then
    MsgBox "Please Enter S,I or R Only", vbOKOnly
    txt_1.Text = ""
    txt_1.SetFocus
Else
   grid_dmbanti.TextMatrix(grid_dmbanti.row, 1) = UCase(Trim(txt_1.Text))
End If
End Sub

Private Sub txt_2_Change()
If Len(Trim(txt_2.Text)) = 1 And Not (UCase(Trim(txt_2.Text)) = "S" Or UCase(Trim(txt_2.Text)) = "I" Or UCase(Trim(txt_2.Text)) = "R") Then
    MsgBox "Please Enter S,I or R Only", vbOKOnly
    txt_2.Text = ""
    txt_2.SetFocus
Else
    grid_dmbanti.TextMatrix(grid_dmbanti.row, 2) = UCase(Trim(txt_2.Text))
End If
End Sub

Private Sub txt_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        '******* Shraddha 30/06/09 ********
        'If Len(TXT_ORGAN3.Text) >= 1 Then
        If Len(cmb_org3.Text) >= 1 Then
        '**********************************
            Call GridControls("VISIBLE", 3)
            Call movecontrols(, 3)
        Else
            
            If grid_dmbanti.row < grid_dmbanti.Rows - 1 Then
                Call GridControls("VISIBLE", 0)
                Call movecontrols("DOWN", 0)
            Else
                GRIDSCROLL = True
                grid_dmbanti.Rows = grid_dmbanti.Rows + 1
                Call GridControls("VISIBLE", 0)
                Call movecontrols("DOWN", 0)
            End If
        End If
        
    End If
    
End Sub

Private Sub txt_1_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        '******** Shrddha 30/06/09 ********
        'If Len(TXT_ORGAN2.Text) >= 1 Then
        If Len(cmb_org2.Text) >= 1 Then
        '**********************************
                'grid_dmbanti.Rows = grid_dmbanti.Rows + 1
                Call GridControls("VISIBLE", 2)
                'Call movecontrols("DOWN", 2)
                Call movecontrols(, 2)
        Else
                If grid_dmbanti.row < grid_dmbanti.Rows - 1 Then
                    Call GridControls("VISIBLE", 0)
                    Call movecontrols("DOWN", 0)
                Else
                    'cmb_antibiotic.clear
                    'GET_ANTIBIOTIC
                    GRIDSCROLL = True
                    grid_dmbanti.Rows = grid_dmbanti.Rows + 1
                    'cmb_antibiotic.Enabled = True
                    Call GridControls("VISIBLE", 0)
                    Call movecontrols("DOWN", 0)
                End If
        End If
    End If

End Sub
 

Private Sub txt_3_Change()
If Len(Trim(txt_3.Text)) = 1 And Not (UCase(Trim(txt_3.Text)) = "S" Or UCase(Trim(txt_3.Text)) = "I" Or UCase(Trim(txt_3.Text)) = "R") Then
    MsgBox "Please Enter S,I or R Only", vbOKOnly
    txt_3.Text = ""
    txt_3.SetFocus
Else
    grid_dmbanti.TextMatrix(grid_dmbanti.row, 3) = UCase(Trim(txt_3.Text))
End If
End Sub

Private Sub txt_3_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

       If grid_dmbanti.row < grid_dmbanti.Rows - 1 Then
              Call GridControls("VISIBLE", 0)
              Call movecontrols("DOWN", 0)
       Else
              GRIDSCROLL = True
              grid_dmbanti.Rows = grid_dmbanti.Rows + 1
              Call GridControls("VISIBLE", 0)
              Call movecontrols("DOWN", 0)
       End If
End If

End Sub

Private Sub txt_bresult_GotFocus()
On Error GoTo m14:
w.Visible = False
Set d = w.ActiveDocument
d.Content.Text = ""
d.Content.Text = txt_bresult.Text
w.Caption = "Full Report"
w.WindowState = wdWindowStateMaximize
d.Content.Text = txt_bresult.Text
w.Visible = True
w.Activate
m14:
On Error Resume Next:
Set w = CreateObject("word.Application")
w.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
w.Caption = "Full Report"
'w.WindowState = wdWindowStateNormal
w.WindowState = wdWindowStateMaximize
Set d = w.ActiveDocument
d.Content.Text = txt_bresult.Text
w.Visible = True
w.Activate
End Sub

Private Sub txt_bresult_LostFocus()
On Error GoTo m15:
Set d = w.ActiveDocument
txt_bresult.Text = d.Content.Text
'flagimp = False
d.Content.Text = ""
' Modified By Pankaj Baggal 03/09/2004
'd.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
d.Close (wdDoNotSaveChanges)
w.quit
'W1.Visible = False
Exit Sub
m15:
MsgBox Err.Description
End Sub

''***** Shraddha 10/05/11 *****
'Private Sub txt_c_1_Change()
'If Len(Trim(txt_c_1.Text)) = 1 And Not (UCase(Trim(txt_c_1.Text)) = "S" Or UCase(Trim(txt_c_1.Text)) = "I" Or UCase(Trim(txt_c_1.Text)) = "R") Then
'    MsgBox "Please Enter S,I or R Only", vbOKOnly
'    txt_c_1.Text = ""
'    txt_c_1.SetFocus
'Else
'    'grid_c_anti.TextMatrix(grid_c_anti.row, 1) = UCase(Trim(txt_c_1.Text))
'End If
'End Sub




Private Sub txt_c_1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

If Len(Trim(grid_c_anti.TextMatrix(grid_c_anti.row, 0))) > 0 Then
    'If Len(Trim(txt_c_1.Text)) = 1 And Not (UCase(Trim(txt_c_1.Text)) = "S" Or UCase(Trim(txt_c_1.Text)) = "R") Then
    If Len(Trim(txt_c_1.Text)) > 1 And Not (UCase(Trim(txt_c_1.Text)) = "S" Or UCase(Trim(txt_c_1.Text)) = "R") Then
        MsgBox "Please Enter S or R Only", vbOKOnly
        txt_c_1.Text = ""
        txt_c_1.SetFocus
        Exit Sub
    Else
       grid_c_anti.TextMatrix(grid_c_anti.row, 1) = UCase(Trim(txt_c_1.Text))
       If grid_c_anti.row < grid_c_anti.Rows - 1 Then
            Call GridControls_mb("VISIBLE", 0)
            Call movecontrols_MB("DOWN", 0)
       Else
            GRIDSCROLL = True
            grid_c_anti.Rows = grid_c_anti.Rows + 1
            Call GridControls_mb("VISIBLE", 0)
            Call movecontrols_MB("DOWN", 0)
       End If
    End If
Else
    MsgBox ("Please select antibiotic first....!"), vbOKOnly
End If
End If
End Sub

'********* 10/05/11 **********

'
'txt_1.Text = Trim(UCase(txt_1.Text))
'If Len(Trim(txt_1.Text)) = 1 And Not (txt_1.Text = "S" Or txt_1.Text = "I" Or txt_1.Text = "R") Then
'    MsgBox "Please Enter S,I or R Only", vbOKOnly
'    txt_1.Text = ""
'    Exit Sub
'ElseIf Len(txt_1.Text) = 0 Then
'    grid_anti.TextMatrix(grid_anti.row, 1) = txt_1.Text
'ElseIf Len(Trim(TXT_ORGAN2.Text)) = 1 Then
'    grid_anti.TextMatrix(grid_anti.row, 1) = txt_1.Text
'    txt_2.Text = CStr(grid_anti.TextMatrix(grid_anti.row, 2))
'    txt_2.SetFocus
'Else
'     grid_anti.TextMatrix(grid_anti.row, 1) = txt_1.Text
'End If
'
'
'   If KeyAscii = 9 Then
'
'    If grid_anti.Rows - 1 = grid_anti.row Then
'
'            If Len(Trim(txt_organ1.Text)) >= 1 And Len(Trim(TXT_ORGAN2.Text)) = 0 And Len(Trim(TXT_ORGAN3.Text)) = 0 Then
'                txt_comment.SetFocus
'            End If
'
'    Else
'
'        If Len(Trim(txt_organ1.Text)) >= 1 And Len(Trim(TXT_ORGAN2.Text)) = 0 And Len(Trim(TXT_ORGAN3.Text)) = 0 Then
'            grid_anti.row = grid_anti.row + 1
'            'check for next value and assign to textfield
'
'            txt_1.Text = CStr(grid_anti.TextMatrix(grid_anti.row, 1))
'            'reassign the positioning
'            grid_anti.Col = 1
'            txt_1.Left = grid_anti.CellLeft + grid_anti.Left
'            txt_1.Top = grid_anti.CellTop + grid_anti.Top
'            txt_1.SetFocus
'        End If
'
'    End If
'    End If
'    'If KeyAscii = 13 Then
'     '   txt_comment.SetFocus
'    'End If
'
'End Sub



'Private Sub txt_1_KeyPress(KeyAscii As Integer)
'    txt_1.Text = Trim(UCase(txt_1.Text))
'
'    If Len(Trim(txt_1.Text)) = 1 And Not (txt_1.Text = "S" Or txt_1.Text = "I" Or txt_1.Text = "R") Then
'        MsgBox "Please Enter S,I or R Only", vbOKOnly
'        txt_1.Text = ""
'    Exit Sub
''    ElseIf Len(txt_1.Text) = 0 Then
''        grid_anti.TextMatrix(grid_anti.row, 1) = txt_1.Text
''    ElseIf Len(Trim(TXT_ORGAN2.Text)) = 1 Then
''        grid_anti.TextMatrix(grid_anti.row, 1) = txt_1.Text
''        txt_2.Text = CStr(grid_anti.TextMatrix(grid_anti.row, 2))
''        txt_2.SetFocus
'    Else
'         grid_anti.TextMatrix(grid_anti.row, 1) = txt_1.Text
'    End If
'
'    If KeyAscii = 13 Then
'
'        'If Len(Trim(txt_organ1.Text)) >= 1 And Len(Trim(TXT_ORGAN2.Text)) = 0 And Len(Trim(TXT_ORGAN3.Text)) = 0 Then
'
'                'check for next value and assign to textfield
'
'                If grid_anti.row = grid_anti.Rows - 1 Then
'                     grid_anti.row = grid_anti.row
'                Else
'                    grid_anti.row = grid_anti.row + 1
'                End If
'
'                txt_1.Text = CStr(grid_anti.TextMatrix(grid_anti.row, 1))
'                'reassign the positioning
'                grid_anti.Col = 1
'                txt_1.Left = grid_anti.CellLeft + grid_anti.Left
'                txt_1.Top = grid_anti.CellTop + grid_anti.Top
'                'txt_1.SetFocus
'        'End If
'    End If
'
'End Sub
'
'Private Sub txt_2_KeyPress(KeyAscii As Integer)
'
'   txt_2.Text = Trim(UCase(txt_2.Text))
'
'    If Len(Trim(txt_2.Text)) = 1 And Not (txt_2.Text = "S" Or txt_2.Text = "I" Or txt_2.Text = "R") Then
'        MsgBox "Please Enter S,I or R Only", vbOKOnly
'         txt_2.Text = ""
'    Exit Sub
'
''    ElseIf Len(txt_2.Text) = 0 Then
''        grid_anti.TextMatrix(grid_anti.row, 2) = txt_2.Text
''    ElseIf Len(Trim(TXT_ORGAN3.Text)) = 1 Then
''        grid_anti.TextMatrix(grid_anti.row, 2) = txt_2.Text
''        txt_3.Text = CStr(grid_anti.TextMatrix(grid_anti.row, 3))
''        txt_3.SetFocus
'    Else
'         grid_anti.TextMatrix(grid_anti.row, 2) = txt_2.Text
'    End If
'    'txt_2.SetFocus
'
'If KeyAscii = 13 Then
'
'                If grid_anti.row = grid_anti.Rows - 1 Then
'                     grid_anti.row = grid_anti.row
'                Else
'                    grid_anti.row = grid_anti.row + 1
'                End If
'
'                txt_2.Text = CStr(grid_anti.TextMatrix(grid_anti.row, 2))
'                'reassign the positioning
'                grid_anti.Col = 2
'                txt_2.Left = grid_anti.CellLeft + grid_anti.Left
'                txt_2.Top = grid_anti.CellTop + grid_anti.Top
'End If
'
'End Sub
'
'Private Sub txt_3_KeyPress(KeyAscii As Integer)
' txt_3.Text = Trim(UCase(txt_3.Text))
'
'If Len(Trim(txt_3.Text)) = 1 And Not (txt_3.Text = "S" Or txt_3.Text = "I" Or txt_3.Text = "R") Then
'    MsgBox "Please Enter S,I or R Only", vbOKOnly
'    txt_3.Text = ""
'    Exit Sub
''ElseIf Len(Trim(txt_3.Text)) = 0 Then
' '   grid_anti.TextMatrix(grid_anti.row, 3) = txt_3.Text
'Else
'    grid_anti.TextMatrix(grid_anti.row, 3) = txt_3.Text
'End If
'
''    i = grid_anti.Rows - 1
'
'
' '   If i = grid_anti.row Then
'  '      Exit Sub
'    If KeyAscii = 13 Then
'
'                If grid_anti.row = grid_anti.Rows - 1 Then
'                     grid_anti.row = grid_anti.row
'                Else
'                    grid_anti.row = grid_anti.row + 1
'                End If
'        'check for next value and assign to textfield
'
'    '    txt_1.Text = CStr(grid_anti.TextMatrix(grid_anti.row, 1))
'     '   txt_2.Text = CStr(grid_anti.TextMatrix(grid_anti.row, 2))
'        txt_3.Text = CStr(grid_anti.TextMatrix(grid_anti.row, 3))
'        'reassign the positioning
'      '  grid_anti.Col = 1
'       ' txt_1.Left = grid_anti.CellLeft + grid_anti.Left
'       ' txt_1.Top = grid_anti.CellTop + grid_anti.Top
'       ' grid_anti.Col = 2
'       ' txt_2.Left = grid_anti.CellLeft + grid_anti.Left
'       ' txt_2.Top = grid_anti.CellTop + grid_anti.Top
'        grid_anti.Col = 3
'        txt_3.Left = grid_anti.CellLeft + grid_anti.Left
'        txt_3.Top = grid_anti.CellTop + grid_anti.Top
'        txt_3.SetFocus
'    End If
' 'End If
'
'
'End Sub


'Private Sub txt_1_lostfocus()
'txt_1.Text = Trim(UCase(txt_1.Text))
'If Len(Trim(txt_1.Text)) = 1 And Not (txt_1.Text = "S" Or txt_1.Text = "I" Or txt_1.Text = "R") Then
'    MsgBox "Please Enter S,I or R Only", vbOKOnly
'    txt_1.Text = ""
'    Exit Sub
'ElseIf Len(txt_1.Text) = 0 Then
'    grid_anti.TextMatrix(grid_anti.row, 1) = txt_1.Text
'ElseIf Len(Trim(TXT_ORGAN2.Text)) = 1 Then
'    grid_anti.TextMatrix(grid_anti.row, 1) = txt_1.Text
'    txt_2.Text = CStr(grid_anti.TextMatrix(grid_anti.row, 2))
'    txt_2.SetFocus
'Else
'     grid_anti.TextMatrix(grid_anti.row, 1) = txt_1.Text
'End If
'
'If grid_anti.Rows - 1 = grid_anti.row Then
'
'        If Len(Trim(txt_organ1.Text)) >= 1 And Len(Trim(TXT_ORGAN2.Text)) = 0 And Len(Trim(TXT_ORGAN3.Text)) = 0 Then
'            txt_comment.SetFocus
'        End If
'
'        'If Len(Trim(txt_organ1.Text)) >= 1 And Len(Trim(TXT_ORGAN2.Text)) = 0 And Len(Trim(TXT_ORGAN3.Text)) = 0 Then
'         '   txt_comment.SetFocus
'        'End If
'Else
'
'If Len(Trim(txt_organ1.Text)) >= 1 And Len(Trim(TXT_ORGAN2.Text)) = 0 And Len(Trim(TXT_ORGAN3.Text)) = 0 Then
'        grid_anti.row = grid_anti.row + 1
'        'check for next value and assign to textfield
'
'
'        txt_1.Text = CStr(grid_anti.TextMatrix(grid_anti.row, 1))
'        'reassign the positioning
'        grid_anti.Col = 1
'        txt_1.Left = grid_anti.CellLeft + grid_anti.Left
'        txt_1.Top = grid_anti.CellTop + grid_anti.Top
'        'txt_1.SetFocus
'End If
'
'End If
'End Sub


'Private Sub txt_2_lostfocus()
'    'If Trim(txt_2.Text) <> "" Then
'    '   Exit Sub
'    'End If
'   txt_2.Text = Trim(UCase(txt_2.Text))
'If Len(Trim(txt_2.Text)) = 1 And Not (txt_2.Text = "S" Or txt_2.Text = "I" Or txt_2.Text = "R") Then
'    MsgBox "Please Enter S,I or R Only", vbOKOnly
'    txt_2.Text = ""
'    Exit Sub
''ElseIf Len(txt_2.Text) = 0 Then
' '   grid_anti.TextMatrix(grid_anti.row, 2) = txt_2.Text
'ElseIf Len(Trim(TXT_ORGAN3.Text)) >= 1 Then
'    grid_anti.TextMatrix(grid_anti.row, 2) = txt_2.Text
'    txt_3.Text = CStr(grid_anti.TextMatrix(grid_anti.row, 3))
'    txt_3.SetFocus
'Else
'     grid_anti.TextMatrix(grid_anti.row, 2) = txt_2.Text
'
'
''End If
'   i = grid_anti.Rows - 1
'    If i = grid_anti.row Then
'        Exit Sub
'    Else
'
'
'
''If Len(Trim(txt_organ1.Text)) >= 1 And Len(Trim(TXT_ORGAN2.Text)) >= 1 And Len(Trim(TXT_ORGAN3.Text)) = 0 Then
'
'         grid_anti.row = grid_anti.row + 1
'        'check for next value and assign to textfield
'        txt_1.Text = CStr(grid_anti.TextMatrix(grid_anti.row, 1))
'        txt_2.Text = CStr(grid_anti.TextMatrix(grid_anti.row, 2))
'
'        'reassign the positioning
'        grid_anti.Col = 1
'        txt_1.Left = grid_anti.CellLeft + grid_anti.Left
'        txt_1.Top = grid_anti.CellTop + grid_anti.Top
'        grid_anti.Col = 2
'        txt_2.Left = grid_anti.CellLeft + grid_anti.Left
'        txt_2.Top = grid_anti.CellTop + grid_anti.Top
'
'    txt_1.SetFocus
'
'    End If
'End If
'End Sub

'Private Sub txt_3_lostfocus()
'
'    txt_3.Text = Trim(UCase(txt_3.Text))
'
'If Len(Trim(txt_3.Text)) = 1 And Not (txt_3.Text = "S" Or txt_3.Text = "I" Or txt_3.Text = "R") Then
'    MsgBox "Please Enter S,I or R Only", vbOKOnly
'    txt_3.Text = ""
'    Exit Sub
''ElseIf Len(Trim(txt_3.Text)) = 0 Then
'  '  grid_anti.TextMatrix(grid_anti.row, 3) = txt_3.Text
'Else
'    grid_anti.TextMatrix(grid_anti.row, 3) = txt_3.Text
'    i = grid_anti.Rows - 1
'    If i = grid_anti.row Then
'        Exit Sub
'    Else
'        grid_anti.row = grid_anti.row + 1
'        'check for next value and assign to textfield
'        txt_1.Text = CStr(grid_anti.TextMatrix(grid_anti.row, 1))
'        txt_2.Text = CStr(grid_anti.TextMatrix(grid_anti.row, 2))
'        txt_3.Text = CStr(grid_anti.TextMatrix(grid_anti.row, 3))
'        'reassign the positioning
'        grid_anti.Col = 1
'        txt_1.Left = grid_anti.CellLeft + grid_anti.Left
'        txt_1.Top = grid_anti.CellTop + grid_anti.Top
'        grid_anti.Col = 2
'        txt_2.Left = grid_anti.CellLeft + grid_anti.Left
'        txt_2.Top = grid_anti.CellTop + grid_anti.Top
'        grid_anti.Col = 3
'        txt_3.Left = grid_anti.CellLeft + grid_anti.Left
'        txt_3.Top = grid_anti.CellTop + grid_anti.Top
'        'txt_3.SetFocus
'        txt_1.SetFocus
'    End If
' End If
'
'
'End Sub

'''''''''''''''''''''''''''''''''''''''''''''varinder
Private Sub txt_addendum_GotFocus()

    On Error GoTo CreateWord

    If w_grossdes Is Nothing Then GoTo CreateWord

    w_grossdes.Visible = False

    Set d_grossdesc = w_grossdes.ActiveDocument

    d_grossdesc.Content.Text = txt_addendum.Text

    w_grossdes.Caption = "Addendum"
    w_grossdes.WindowState = wdWindowStateMaximize
    w_grossdes.Visible = True
    w_grossdes.Activate

    Exit Sub

CreateWord:

    On Error Resume Next

    Set w_grossdes = CreateObject("Word.Application")

    If w_grossdes Is Nothing Then
        MsgBox "Microsoft Word could not be started."
        Exit Sub
    End If

    w_grossdes.Documents.add "C:\DISREPORTSEC\REPORT.dot"

    Set d_grossdesc = w_grossdes.ActiveDocument

    d_grossdesc.Content.Text = txt_addendum.Text

    w_grossdes.Caption = "Addendum"
    w_grossdes.WindowState = wdWindowStateMaximize
    w_grossdes.Visible = True
    w_grossdes.Activate

End Sub


'Private Sub txt_addendum_GotFocus()
'On Error GoTo m3:
'W_grossdes.Visible = False
'Set d_grossdesc = W_grossdes.ActiveDocument
'd_grossdesc.Content = ""
'd_grossdesc.Content = txt_addendum
'W_grossdes.Caption = "Addendum"
'W_grossdes.WindowState = wdWindowStateMaximize
'd_grossdesc.Content = txt_addendum
'W_grossdes.Visible = True
'W_grossdes.Activate
'm3:
'On Error Resume Next:
'Set W_grossdes = CreateObject("word.Application")
'W_grossdes.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
'W_grossdes.Caption = "Addendum"
'W_grossdes.WindowState = wdWindowStateMaximize
'Set d_grossdesc = W_grossdes.ActiveDocument
'd_grossdesc.Content = txt_addendum.Text
'W_grossdes.Visible = True
'W_grossdes.Activate
'End Sub
'
'Private Sub txt_addendum_LostFocus()
'On Error GoTo m4:
'Set d_grossdesc = W_grossdes.ActiveDocument
'txt_addendum.Text = d_grossdesc.Content
''flagimp = False
'd_grossdesc.Content = ""
'' Modified By Pankaj Baggal 03/09/2004
''d_grossdesc.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
'd_grossdesc.Close (wdDoNotSaveChanges)
'W_grossdes.quit
''W1.Visible = False
'Exit Sub
'm4:
'MsgBox Err.Description
'End Sub

Private Sub txt_c_1_LOSTFOCUS()
txt_c_1.Text = Trim(UCase(txt_c_1.Text))
If Len(txt_c_1.Text) = 1 And Not (txt_c_1.Text = "S" Or txt_c_1.Text = "R") Then
    MsgBox "Please Enter S or R Only", vbOKOnly
    txt_c_1.Text = ""
    txt_c_1.SetFocus
    Exit Sub
'ElseIf Len(txt_c_1.Text) = 0 Then
'    grid_c_anti.TextMatrix(grid_c_anti.row, 1) = txt_c_1.Text
'Else
'    grid_c_anti.TextMatrix(grid_c_anti.row, 1) = txt_c_1.Text
'    i = grid_c_anti.Rows - 1
'    If i = grid_c_anti.row Then
'        txt_c_1.Visible = False
'        grid_dmbanti.TextMatrix(grid_dmbanti.row, 3) = UCase(Trim(txt_3.Text))
'        Exit Sub
'    Else
'        grid_c_anti.row = grid_c_anti.row + 1
'
'        txt_c_1.Text = CStr(grid_c_anti.TextMatrix(grid_c_anti.row, 1))
'        txt_c_1.Left = grid_c_anti.CellLeft + grid_c_anti.Left
'        txt_c_1.Top = grid_c_anti.CellTop + grid_c_anti.Top
'        txt_c_1.SetFocus
'     End If
End If



End Sub
'************ Added and commented by Shraddha 02/07/09 ******************
Private Sub cmb_org3_lostfocus()
'On Error GoTo MSG
 
'**************** Shraddha 03/07/09 **************
Dim o3, m3
Dim n3, l3 As Integer
l3 = Len(cmb_org3.Text)
n3 = InStr(1, cmb_org3.Text, "$")

If n3 <> 0 Then
    m3 = Mid(cmb_org3, n3 + 1, l3 - n3)
    o3 = Mid(cmb_org3, 1, n3 - 1)
    
    If m3 = "Heavy" Then
        cmb_org3.Text = "Heavy growth of " & o3
    End If
    
    If m3 = "Poor" Then
        cmb_org3.Text = "Poor growth of " & o3
    End If
    
    If m3 = "Moderate" Then
        cmb_org3.Text = "Moderate growth of " & o3
    End If
    
    If m3 = "Significant" Then
        cmb_org3.Text = "Significant growth of " & o3
    End If
End If
'**************************************************
 
GET_ANTIBIOTIC

If new_flag Then
  
  '***** Shraddha 30/06/09 *****
  'If Len(Trim(txt_organ1.Text)) >= 1 Then
    If Len(Trim(cmb_org1.Text)) >= 1 Then
  
        'If Len(Trim(TXT_ORGAN2.Text)) >= 1 Then
        If Len(Trim(cmb_org2.Text)) >= 1 Then
        
            'If Len(Trim(TXT_ORGAN3.Text)) >= 1 Then
            If Len(Trim(cmb_org3.Text)) >= 1 Then
  '*****************************
            
                ''''THREE ORG
                grid_dmbanti.Rows = 2
                grid_dmbanti.Cols = 4
                grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           |^         2           |^         3           "
                
            Else
                '''TWO ORG
                grid_dmbanti.Rows = 2
                grid_dmbanti.Cols = 3
                grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           |^         2           "
                
            End If
        Else
                '''ONE ORG
                grid_dmbanti.Rows = 2
                grid_dmbanti.Cols = 2
                grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           "
                     
        End If
  Else
            
  End If
  
  Call movecontrols(, 0)
  Call GridControls("VISIBLE", 0)

ElseIf VIEW_FLAG Then

ElseIf mod_flag Then

    '***** Shraddha 30/06/09 *****
    'If Len(Trim(txt_organ1.Text)) >= 1 Then
    If Len(Trim(cmb_org1.Text)) >= 1 Then
  
        'If Len(Trim(TXT_ORGAN2.Text)) >= 1 Then
        If Len(Trim(cmb_org2.Text)) >= 1 Then
        
            'If Len(Trim(TXT_ORGAN3.Text)) >= 1 Then
            If Len(Trim(cmb_org3.Text)) >= 1 Then
    '*******************************
                ''''THREE ORG
                grid_dmbanti.Rows = 1
                grid_dmbanti.Cols = 4
                grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           |^         2           |^         3           "

            Else
                '''TWO ORG
                grid_dmbanti.Rows = 1
                grid_dmbanti.Cols = 3
                grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           |^         2           "

            End If
        Else
                '''ONE ORG
                grid_dmbanti.Rows = 1
                grid_dmbanti.Cols = 2
                grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           "
                              
        End If
    End If
        
''getresult_bacteriology
        sel_forupd2_qr(0) = mod_reqno
        Set sel_forupd2_rs = sel_forupd2_qr.OpenResultset(1, 3)
        
        If sel_forupd2_rs.RowCount > 0 Then
                While Not sel_forupd2_rs.EOF
                    If grid_dmbanti.Cols = 2 Then
                            While Not sel_forupd2_rs.EOF
                                grid_dmbanti.AddItem sel_forupd2_rs("antibiotic") & vbTab & sel_forupd2_rs("antibio1")
                                sel_forupd2_rs.MoveNext
                            Wend
                    ElseIf grid_dmbanti.Cols = 3 Then
                            While Not sel_forupd2_rs.EOF
                                grid_dmbanti.AddItem sel_forupd2_rs("antibiotic") & vbTab & sel_forupd2_rs("antibio1") & vbTab & sel_forupd2_rs("antibio2")
                                sel_forupd2_rs.MoveNext
                            Wend
                    ElseIf grid_dmbanti.Cols = 4 Then
                            While Not sel_forupd2_rs.EOF
                                grid_dmbanti.AddItem sel_forupd2_rs("antibiotic") & vbTab & sel_forupd2_rs("antibio1") & vbTab & sel_forupd2_rs("antibio2") & vbTab & sel_forupd2_rs("antibio3")
                                sel_forupd2_rs.MoveNext
                            Wend
                    End If
                     
                Wend
        Else
            
            grid_dmbanti.Rows = grid_dmbanti.Rows + 1
            Call movecontrols(, 0)
            Call GridControls("VISIBLE", 0)
            
        End If
   'cmb_antibiotic.SetFocus 'modified by mrunal 26_2
End If
'cmb_org2.SetFocus

    Exit Sub
MSG:
    
    MsgBox ERROR(Err.Description)
End Sub


'Private Sub txt_organ3_lostfocus()
'On Error GoTo MSG
'
'
' GET_ANTIBIOTIC
'
'If new_flag Then
'
'  '***** Shraddha 30/06/09 *****
'  'If Len(Trim(txt_organ1.Text)) >= 1 Then
'  If Len(Trim(cmb_org1.Text)) >= 1 Then
'
'        'If Len(Trim(TXT_ORGAN2.Text)) >= 1 Then
'        If Len(Trim(cmb_org2.Text)) >= 1 Then
'
'            'If Len(Trim(TXT_ORGAN3.Text)) >= 1 Then
'            If Len(Trim(cmb_org3.Text)) >= 1 Then
'  '*****************************
'
'                ''''THREE ORG
'                grid_dmbanti.Rows = 2
'                grid_dmbanti.Cols = 4
'                grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           |^         2           |^         3           "
'
'            Else
'                '''TWO ORG
'                grid_dmbanti.Rows = 2
'                grid_dmbanti.Cols = 3
'                grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           |^         2           "
'
'            End If
'        Else
'                '''ONE ORG
'                grid_dmbanti.Rows = 2
'                grid_dmbanti.Cols = 2
'                grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           "
'
'        End If
'  Else
'
'        'MsgBox ("Atleast 1 Organism is needed......! "), vbOKOnly
'        'txt_organ1.SetFocus
'
'        'Exit Sub
'
'  End If
'
'  Call movecontrols(, 0)
'  Call GridControls("VISIBLE", 0)
'
'
'
'ElseIf VIEW_FLAG Then
'
'
'ElseIf mod_flag Then
'
'    '***** Shraddha 30/06/09 *****
'    'If Len(Trim(txt_organ1.Text)) >= 1 Then
'    If Len(Trim(cmb_org1.Text)) >= 1 Then
'
'        'If Len(Trim(TXT_ORGAN2.Text)) >= 1 Then
'        If Len(Trim(cmb_org2.Text)) >= 1 Then
'
'            'If Len(Trim(TXT_ORGAN3.Text)) >= 1 Then
'            If Len(Trim(cmb_org3.Text)) >= 1 Then
'    '*******************************
'                ''''THREE ORG
'                grid_dmbanti.Rows = 1
'                grid_dmbanti.Cols = 4
'                grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           |^         2           |^         3           "
'
'            Else
'                '''TWO ORG
'                grid_dmbanti.Rows = 1
'                grid_dmbanti.Cols = 3
'                grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           |^         2           "
'
'            End If
'        Else
'                '''ONE ORG
'                grid_dmbanti.Rows = 1
'                grid_dmbanti.Cols = 2
'                grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           "
'
'        End If
'    End If
'
'        'grid_dmbanti.clear
'''getresult_bacteriology
'        sel_forupd2_qr(0) = mod_reqno
'        Set sel_forupd2_rs = sel_forupd2_qr.OpenResultset(1, 3)
'
'        If sel_forupd2_rs.RowCount > 0 Then
'                While Not sel_forupd2_rs.EOF
'                    If grid_dmbanti.Cols = 2 Then
'                            While Not sel_forupd2_rs.EOF
'                                grid_dmbanti.AddItem sel_forupd2_rs("antibiotic") & vbTab & sel_forupd2_rs("antibio1")
'                                sel_forupd2_rs.MoveNext
'                            Wend
'                    ElseIf grid_dmbanti.Cols = 3 Then
'                            While Not sel_forupd2_rs.EOF
'                                grid_dmbanti.AddItem sel_forupd2_rs("antibiotic") & vbTab & sel_forupd2_rs("antibio1") & vbTab & sel_forupd2_rs("antibio2")
'                                sel_forupd2_rs.MoveNext
'                            Wend
'                    ElseIf grid_dmbanti.Cols = 4 Then
'                            While Not sel_forupd2_rs.EOF
'                                grid_dmbanti.AddItem sel_forupd2_rs("antibiotic") & vbTab & sel_forupd2_rs("antibio1") & vbTab & sel_forupd2_rs("antibio2") & vbTab & sel_forupd2_rs("antibio3")
'                                sel_forupd2_rs.MoveNext
'                            Wend
'                    End If
'
'                Wend
'        Else
'
'            grid_dmbanti.Rows = grid_dmbanti.Rows + 1
'            Call movecontrols(, 0)
'            Call GridControls("VISIBLE", 0)
'
'        End If
'        'Call movecontrols(, 0)
'        'Call GridControls("VISIBLE", 0)
'
'End If
'
'
' '''commented by kanchan  22/04/2008
'  '  LOAD_ANTIBIOTIC
'    Exit Sub
'MSG:
'
'    MsgBox Error(Err.Description)
'End Sub
'************************************************************
Private Sub GridControls(MODE As String, Optional Cl As Integer)
    
''''' Done by Kanchan for proper movement of controls on the grid

    If MODE = "INVISIBLE" Then
           
        cmb_antibiotic.Visible = False
        txt_1.Visible = False
        txt_2.Visible = False
        txt_3.Visible = False
                       
    ElseIf MODE = "VISIBLE" Then
        cmb_antibiotic.Visible = False
        txt_1.Visible = False
        txt_2.Visible = False
        txt_3.Visible = False
                
        Select Case Cl
    
        Case 0
            cmb_antibiotic.Visible = True
            cmb_antibiotic.SetFocus
        
        Case 1
            txt_1.Visible = True
            txt_1.SetFocus
        
        Case 2
            txt_2.Visible = True
            txt_2.SetFocus
        
        Case 3
            txt_3.Visible = True
            txt_3.SetFocus
        End Select
    End If
End Sub
Private Sub movecontrols(Optional Direction As String, Optional Cl As Integer)
    
'''' Done by Kanchan for proper movement of controls on the grid
        
    With grid_dmbanti
        Select Case Direction
            Case "DOWN"
                If .row < .Rows - 1 Then
                    .row = .row + 1
                End If
            Case "UP"
                If .row > 1 Then
                    .row = .row - 1
                End If
            End Select
         Select Case Cl
         Case 0
            .col = 0
             cmb_antibiotic.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth
             cmb_antibiotic.Text = .TextMatrix(.row, 0)
             cmb_antibiotic.Visible = True
             cmb_antibiotic.SetFocus
             
        Case 1
            .col = 1
            txt_1.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth
            txt_1.Text = .TextMatrix(.row, 1)
            txt_1.Visible = True
            txt_1.SetFocus
        
        Case 2
            .col = 2
            txt_2.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth
            txt_2.Text = .TextMatrix(.row, 2)
            txt_2.Visible = True
'            txt_2.SetFocus
            
         Case 3
            .col = 3
            txt_3.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth
            txt_3.Text = .TextMatrix(.row, 3)
            txt_3.Visible = True
            txt_3.SetFocus
        End Select
      End With

End Sub

Private Sub movecontrols_MB(Optional Direction As String, Optional Cl As Integer)
With grid_c_anti
    Select Case Direction
        Case "DOWN"
            If .row < .Rows - 1 Then
                .row = .row + 1
            End If
        Case "UP"
            If .row > 1 Then
                .row = .row - 1
            End If
    End Select
        
    Select Case Cl
        Case 0
            .col = 0
             cmb_antiboitic_mb.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth
             cmb_antiboitic_mb.Text = .TextMatrix(.row, 0)
             cmb_antiboitic_mb.Visible = True
             cmb_antiboitic_mb.SetFocus
        Case 1
            .col = 1
            txt_c_1.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth
            txt_c_1.Text = .TextMatrix(.row, 1)
            txt_c_1.Visible = True
            txt_c_1.SetFocus
'        Case 2
'            .col = 2
'            txt_2.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth
'            txt_2.Text = .TextMatrix(.row, 2)
'            txt_2.Visible = True
'        Case 3
'            .col = 3
'            txt_3.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth
'            txt_3.Text = .TextMatrix(.row, 3)
'            txt_3.Visible = True
'            txt_3.SetFocus
    End Select
End With
End Sub

Private Sub GridControls_mb(MODE As String, Optional Cl As Integer)
    
''''' Done by Kanchan for proper movement of controls on the grid

    If MODE = "INVISIBLE" Then
           
        cmb_antiboitic_mb.Visible = False
        txt_c_1.Visible = False
'        txt_2.Visible = False
'        txt_3.Visible = False
                       
    ElseIf MODE = "VISIBLE" Then
        cmb_antiboitic_mb.Visible = True
        If new_flag = True Then
            txt_c_1.Visible = True
        Else
            txt_c_1.Visible = False
        End If
'        txt_2.Visible = False
'        txt_3.Visible = False
                
        Select Case Cl
    
        Case 0
            cmb_antiboitic_mb.Visible = True
            cmb_antiboitic_mb.SetFocus
        
        Case 1
            txt_c_1.Visible = True
            txt_c_1.SetFocus
        
'        Case 2
'            txt_2.Visible = True
'            txt_2.SetFocus
'
'        Case 3
'            txt_3.Visible = True
'            txt_3.SetFocus
        End Select
    End If
End Sub

Private Sub GET_ANTIBIOTIC()
On Error GoTo MSG
cmb_antibiotic.clear
Dim RCN As Integer
Dim CHKFLG As Boolean

    If new_flag Then
       
            sel_antiqr(0) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
            sel_antiqr(1) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
            Set sel_antirs = sel_antiqr.OpenResultset(1, 3)
            'If Len(TXT_ORGAN2.Text) >= 1 Then
             '   If Len(TXT_ORGAN3.Text) >= 1 Then
                'three organisms
             '   grid_anti.Rows = 1
             '   grid_anti.Cols = 2
             '   grid_anti.FormatString = "    ANTIBIOTIC                               |^         1           |^          2         |^              3      "
             '   Dim a
                
                While Not sel_antirs.EOF
                    cmb_antibiotic.AddItem sel_antirs("antibdesc")
                    'grid_anti.AddItem sel_antirs("antibdesc")
                    sel_antirs.MoveNext
                Wend
              
       
    End If
    If VIEW_FLAG Then
        CHKFLG = False
            sel_antiqr(0) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
            sel_antiqr(1) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
            Set sel_antirs = sel_antiqr.OpenResultset(1, 3)
                                      
                While Not sel_antirs.EOF
                    CHKFLG = False
                    
                   For RCN = 1 To grid_dmbanti.Rows - 1
                    'StrComp
                    If Trim(grid_dmbanti.TextMatrix(RCN, 0)) = Trim(sel_antirs("ANTIBDESC")) Then
                        
                        CHKFLG = True
                    End If
                   Next RCN
                   
                        If CHKFLG = False Then
                        cmb_antibiotic.AddItem sel_antirs("antibdesc")
                        End If
                   
                  sel_antirs.MoveNext
                Wend

    End If
        
        If mod_flag Then
        CHKFLG = False
            sel_antiqr(0) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
            sel_antiqr(1) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
            Set sel_antirs = sel_antiqr.OpenResultset(1, 3)
       
       
                                   
                While Not sel_antirs.EOF
                    CHKFLG = False
                    
                   For RCN = 1 To grid_dmbanti.Rows - 1
                    'StrComp
                    If Trim(grid_dmbanti.TextMatrix(RCN, 0)) = Trim(sel_antirs("ANTIBDESC")) Then
                        
                        CHKFLG = True
                    End If
                   Next RCN
                   
                        If CHKFLG = False Then
                        cmb_antibiotic.AddItem sel_antirs("antibdesc")
                        End If
                   
                  sel_antirs.MoveNext
                Wend

    End If

    Exit Sub
       

MSG:
    MsgBox ERROR(Err.Description)

End Sub

Private Sub GET_ANTIBIOTIC_MB()
'On Error GoTo MSG
cmb_antiboitic_mb.clear
Dim RCN As Integer
Dim CHKFLG As Boolean
Dim ALLANTIB As Boolean
ALLANTIB = False
    If new_flag Then
        sel_antiqr(0) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
        sel_antiqr(1) = Trim(cmb_organisms.Text)
        Set sel_antirs = sel_antiqr.OpenResultset(1, 3)
        If sel_antirs.EOF = True Then
        ALLANTIB = True
        End If
        
            While Not sel_antirs.EOF
            cmb_antiboitic_mb.AddItem sel_antirs("antibdesc")
            sel_antirs.MoveNext
        Wend
     
    End If
    If VIEW_FLAG Then
        CHKFLG = False
        sel_antiqr(0) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
        sel_antiqr(1) = Trim(cmb_organisms.Text)
            Set sel_antirs = sel_antiqr.OpenResultset(1, 3)
                If sel_antirs.EOF = True Then
                   ALLANTIB = True
                   End If
                While Not sel_antirs.EOF
                    CHKFLG = False
                    
                   For RCN = 1 To grid_c_anti.Rows - 1
                    'StrComp
                    If Trim(grid_c_anti.TextMatrix(RCN, 0)) = Trim(sel_antirs("ANTIBDESC")) Then
                        
                        CHKFLG = True
                    End If
                   Next RCN
                   
                        If CHKFLG = False Then
                        cmb_antiboitic_mb.AddItem sel_antirs("antibdesc")
                        End If
                   
                  sel_antirs.MoveNext
                Wend
        
    End If
        
        If mod_flag Then
        CHKFLG = False
        sel_antiqr(0) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
        sel_antiqr(1) = Trim(cmb_organisms.Text)
            Set sel_antirs = sel_antiqr.OpenResultset(1, 3)
                If sel_antirs.EOF = True Then
                 ALLANTIB = True
                 End If
                While Not sel_antirs.EOF
                    CHKFLG = False
                    
                   For RCN = 1 To grid_c_anti.Rows - 1
                    'StrComp
                    If Trim(grid_c_anti.TextMatrix(RCN, 0)) = Trim(sel_antirs("ANTIBDESC")) Then
                        
                        CHKFLG = True
                    End If
                   Next RCN
                   
                        If CHKFLG = False Then
                        cmb_antiboitic_mb.AddItem sel_antirs("antibdesc")
                        End If
                   
                  sel_antirs.MoveNext
                Wend
       
    End If

    
       If ALLANTIB = True Then
            Dim sel_allantiqr As rdoQuery
            Dim sel_allantirs As rdoResultset
            Dim str1 As String
            Dim S1 As String
            S1 = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
'                str1 = "select antibcode,antibdesc from dislib.dmicbantib WHERE MODALITY= '" & Trim(S1) & "' ORDER BY ANTIBDESC"
                str1 = "select DISTINCT antibdesc from dislib.dmicbantib WHERE MODALITY=? ORDER BY ANTIBDESC"
                Set sel_allantiqr = rdb.CreateQuery("", str1)
                sel_allantiqr(0) = S1
                Set sel_allantirs = sel_allantiqr.OpenResultset(1, 3)
                
                While Not sel_allantirs.EOF
                    CHKFLG = False
                    
                   For RCN = 1 To grid_c_anti.Rows - 1
                    'StrComp
                    If Trim(grid_c_anti.TextMatrix(RCN, 0)) = Trim(sel_allantirs("ANTIBDESC")) Then
                        
                        CHKFLG = True
                    End If
                   Next RCN
                   
                        If CHKFLG = False Then
                        cmb_antiboitic_mb.AddItem sel_allantirs("antibdesc")
                        End If
                   
                  sel_allantirs.MoveNext
                Wend
               
       Else
                                   
       End If
Exit Sub
MSG:
    MsgBox ERROR(Err.Description)

End Sub

Private Sub txt_result_gotfocus()
''On Error GoTo MSG
'    Set objWord = CreateObject("word.application")
'    objWord.Visible = False
'    objWord.WindowState = wdWindowStateMaximize
'    Set objDoc = objWord.Documents.add
'    'objWord.Selection.InsertBefore Text:=("Hi...!!!")
'    objDoc.Content = txt_result.Text
'    objWord.Visible = True
'    Exit Sub
'MSG:
'    'objWord.Repeat
'    If Not Err.Number = -2147023174 Then
'        MsgBox "error" & Err.Number
'    End If
    
End Sub
Private Sub txt_result_lostfocus()
'On Error GoTo MSG
'    txt_result.Text = objDoc.Content
'    'objDoc.Content = ""
'    objDoc.Content = "popopopopo"
'    ' Modified By Pankaj Baggal 03/09/2004
'    'objDoc.SaveAs "c:\windows\system\frozen.doc"
'    objDoc.Close (wdDoNotSaveChanges)
'    objWord.quit
'    Exit Sub
'MSG:
'    If Not Err.Number = -2147023174 Then
'        MsgBox "error" & Err.Number
'    End If
End Sub
Private Sub select_antibiotic()
On Error GoTo MSG

    If new_flag Then
        '********* Shraddha 30/06/09 *********
         'If Len(txt_organ1.Text) >= 1 Then
         If Len(cmb_org1.Text) >= 1 Then
         
            'If Len(TXT_ORGAN2.Text) >= 1 Then
            If Len(cmb_org2.Text) >= 1 Then
            
                'If (TXT_ORGAN3.Text) >= 1 Then
                If (cmb_org3.Text) >= 1 Then
        '************************************
                    ''''three orgnism
                Else
                    ''''two organism
                End If
            Else
                 '''one organism
                 grid_dmbanti.Rows = 2
                grid_dmbanti.Cols = 2
                grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           "
                
                
                grid_dmbanti.row = 1
                grid_dmbanti.col = 0
                cmb_antibiotic.Width = grid_dmbanti.CellWidth
                'cmb_antibiotic.Height = grid_dmbanti.CellHeight
                cmb_antibiotic.Left = grid_dmbanti.CellLeft + grid_dmbanti.Left
                cmb_antibiotic.Top = grid_dmbanti.CellTop + grid_dmbanti.Top
                cmb_antibiotic.Visible = True
                
            End If
         Else
            MsgBox ("Select atleast one organism.....!"), vbOKOnly
         End If
    End If
    
    Exit Sub

MSG:
MsgBox ERROR(Err)
End Sub
Private Sub sel_antibiotic()
On Error GoTo MSG
    If new_flag Then
        '******** Shraddha 30/06/09 ********
        'If Len(txt_organ1.Text) >= 1 Then
        If Len(cmb_org1.Text) >= 1 Then
        
            'If Len(TXT_ORGAN2.Text) >= 1 Then
            If Len(cmb_org2.Text) >= 1 Then
                
                'If Len(TXT_ORGAN3.Text) >= 1 Then
                If Len(cmb_org3.Text) >= 1 Then
        '***********************************
                    grid_dmbanti.Rows = 1
                    grid_dmbanti.Cols = 2
                    grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           |^          2         |^              3      "
                    
                    grid_dmbanti.row = 1
                    grid_dmbanti.col = 1
                    txt_1.Width = grid_dmbanti.CellWidth
                    txt_1.Height = grid_dmbanti.CellHeight
                    txt_1.Left = grid_dmbanti.CellLeft + grid_dmbanti.Left
                    txt_1.Top = grid_dmbanti.CellTop + grid_dmbanti.Top
                    grid_dmbanti.row = 1
                    grid_dmbanti.col = 2
                    txt_2.Width = grid_dmbanti.CellWidth
                    txt_2.Height = grid_dmbanti.CellHeight
                    txt_2.Left = grid_dmbanti.CellLeft + grid_dmbanti.Left
                    txt_2.Top = grid_dmbanti.CellTop + grid_dmbanti.Top
                    grid_dmbanti.row = 1
                    grid_dmbanti.col = 3
                    txt_3.Width = grid_dmbanti.CellWidth
                    txt_3.Height = grid_dmbanti.CellHeight
                    txt_3.Left = grid_dmbanti.CellLeft + grid_dmbanti.Left
                    txt_3.Top = grid_dmbanti.CellTop + grid_dmbanti.Top
                    txt_3.Visible = True
                    txt_2.Visible = True
                    txt_1.Visible = True
            
                Else
                'two organisms
                grid_dmbanti.Rows = 1
                grid_dmbanti.Cols = 2
                grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           |^          2         "
                
                    grid_dmbanti.row = 1
                    grid_dmbanti.col = 1
                    txt_1.Width = grid_dmbanti.CellWidth
                    txt_1.Height = grid_dmbanti.CellHeight
                    txt_1.Left = grid_dmbanti.CellLeft + grid_dmbanti.Left
                    txt_1.Top = grid_dmbanti.CellTop + grid_dmbanti.Top
                    grid_dmbanti.row = 1
                    grid_dmbanti.col = 2
                    txt_2.Width = grid_dmbanti.CellWidth
                    txt_2.Height = grid_dmbanti.CellHeight
                    txt_2.Left = grid_dmbanti.CellLeft + grid_dmbanti.Left
                    txt_2.Top = grid_dmbanti.CellTop + grid_dmbanti.Top
                    txt_3.Visible = False
                    txt_2.Visible = True
                    txt_1.Visible = True
                    txt_1.SetFocus
                End If
            Else
            'only one organisms
            grid_dmbanti.Rows = 2
            grid_dmbanti.Cols = 2
            grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           "
                
                
                grid_dmbanti.row = 1
                grid_dmbanti.col = 0
                cmb_antibiotic.Width = grid_dmbanti.CellWidth
                'cmb_antibiotic.Height = grid_dmbanti.CellHeight
                cmb_antibiotic.Left = grid_dmbanti.CellLeft + grid_dmbanti.Left
                cmb_antibiotic.Top = grid_dmbanti.CellTop + grid_dmbanti.Top
                                
                grid_dmbanti.row = 1
                grid_dmbanti.col = 1
                txt_1.Width = grid_dmbanti.CellWidth
                txt_1.Height = grid_dmbanti.CellHeight
                txt_1.Left = grid_dmbanti.CellLeft + grid_dmbanti.Left
                txt_1.Top = grid_dmbanti.CellTop + grid_dmbanti.Top
                txt_1.Visible = False
                txt_2.Visible = False
                txt_3.Visible = False
                'txt_1.SetFocus
            End If
           Else
            MsgBox "Atleast 1 organisms is needed", vbOKOnly
            Exit Sub
         End If
    End If
    
    Exit Sub
    
MSG:
       MsgBox ERROR(Err)
End Sub
Private Sub LOAD_ANTIBIOTIC()
On Error GoTo MSG
    'check all organisms has been entered or not
    If new_flag Then
        '****** Shraddha 30/06/09 *********
        'If Len(txt_organ1.Text) >= 1 Then
        If Len(cmb_org1.Text) >= 1 Then
        
            'sel_antiqr(0) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
            'sel_antiqr(1) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
            'Set sel_antirs = sel_antiqr.OpenResultset(1, 3)
            
            'If Len(TXT_ORGAN2.Text) >= 1 Then
            If Len(cmb_org2.Text) >= 1 Then
            
                'If Len(TXT_ORGAN3.Text) >= 1 Then
                If Len(cmb_org3.Text) >= 1 Then
            '**********************************
                'three organisms
                grid_anti.Rows = 1
                grid_anti.Cols = 2
                grid_anti.FormatString = "    ANTIBIOTIC                               |^         1           |^          2         |^              3      "
                Dim a
                
                'While Not sel_antirs.EOF
                 '   grid_anti.AddItem sel_antirs("antibdesc")
                 '   sel_antirs.MoveNext
                'Wend
                    grid_anti.row = 1
                    grid_anti.col = 1
                    txt_1.Width = grid_anti.CellWidth
                    txt_1.Height = grid_anti.CellHeight
                    txt_1.Left = grid_anti.CellLeft + grid_anti.Left
                    txt_1.Top = grid_anti.CellTop + grid_anti.Top
                    grid_anti.row = 1
                    grid_anti.col = 2
                    txt_2.Width = grid_anti.CellWidth
                    txt_2.Height = grid_anti.CellHeight
                    txt_2.Left = grid_anti.CellLeft + grid_anti.Left
                    txt_2.Top = grid_anti.CellTop + grid_anti.Top
                    grid_anti.row = 1
                    grid_anti.col = 3
                    txt_3.Width = grid_anti.CellWidth
                    txt_3.Height = grid_anti.CellHeight
                    txt_3.Left = grid_anti.CellLeft + grid_anti.Left
                    txt_3.Top = grid_anti.CellTop + grid_anti.Top
                    txt_3.Visible = True
                    txt_2.Visible = True
                    txt_1.Visible = True
                Else
                'two organisms
                grid_anti.Rows = 1
                grid_anti.Cols = 2
                grid_anti.FormatString = "    ANTIBIOTIC                               |^         1           |^          2         "
                While Not sel_antirs.EOF
                    grid_anti.AddItem sel_antirs("antibdesc")
                    sel_antirs.MoveNext
                Wend
                    grid_anti.row = 1
                    grid_anti.col = 1
                    txt_1.Width = grid_anti.CellWidth
                    txt_1.Height = grid_anti.CellHeight
                    txt_1.Left = grid_anti.CellLeft + grid_anti.Left
                    txt_1.Top = grid_anti.CellTop + grid_anti.Top
                    grid_anti.row = 1
                    grid_anti.col = 2
                    txt_2.Width = grid_anti.CellWidth
                    txt_2.Height = grid_anti.CellHeight
                    txt_2.Left = grid_anti.CellLeft + grid_anti.Left
                    txt_2.Top = grid_anti.CellTop + grid_anti.Top
                    txt_3.Visible = False
                    txt_2.Visible = True
                    txt_1.Visible = True
                    txt_1.SetFocus
                End If
            Else
            'only one organisms
            grid_anti.Rows = 1
            grid_anti.Cols = 2
            grid_anti.FormatString = "    ANTIBIOTIC                               |^         1           "
            While Not sel_antirs.EOF
              grid_anti.AddItem sel_antirs("antibdesc")
              sel_antirs.MoveNext
            Wend
                grid_anti.row = 1
                grid_anti.col = 1
                txt_1.Width = grid_anti.CellWidth
                txt_1.Height = grid_anti.CellHeight
                txt_1.Left = grid_anti.CellLeft + grid_anti.Left
                txt_1.Top = grid_anti.CellTop + grid_anti.Top
                txt_1.Visible = True
                txt_2.Visible = False
                txt_3.Visible = False
                txt_1.SetFocus
            End If
        Else
            MsgBox "Atleast 1 organisms is needed", vbOKOnly
            Exit Sub
         End If
    ElseIf VIEW_FLAG Then
    
            '******* Shraddha 30/06/09 *******
            'If Len(txt_organ1.Text) >= 1 Then
            If Len(cmb_org1.Text) >= 1 Then
            
                 'If Len(TXT_ORGAN2.Text) >= 1 Then
                 If Len(cmb_org2.Text) >= 1 Then
                 
                     'If Len(TXT_ORGAN3.Text) >= 1 Then
                     If Len(cmb_org3.Text) >= 1 Then
            '*********************************
                        'three organisms
                        grid_anti.row = 1
                        grid_anti.col = 1
                        txt_1.Text = grid_anti.TextMatrix(1, 1)
                        txt_1.Width = grid_anti.CellWidth
                        txt_1.Height = grid_anti.CellHeight
                        txt_1.Left = grid_anti.CellLeft + grid_anti.Left
                        txt_1.Top = grid_anti.CellTop + grid_anti.Top
                        txt_1.Visible = True
                        grid_anti.row = 1
                        grid_anti.col = 2
                        txt_2.Text = grid_anti.TextMatrix(1, 2)
                        txt_2.Width = grid_anti.CellWidth
                        txt_2.Height = grid_anti.CellHeight
                        txt_2.Left = grid_anti.CellLeft + grid_anti.Left
                        txt_2.Top = grid_anti.CellTop + grid_anti.Top
                        txt_2.Visible = True
                        grid_anti.row = 1
                        grid_anti.col = 3
                        txt_3.Text = grid_anti.TextMatrix(1, 3)
                        txt_3.Width = grid_anti.CellWidth
                        txt_3.Height = grid_anti.CellHeight
                        txt_3.Left = grid_anti.CellLeft + grid_anti.Left
                        txt_3.Top = grid_anti.CellTop + grid_anti.Top
                        txt_3.Visible = True
                    Else
                        'two organisms
                        grid_anti.Cols = 3
                        grid_anti.row = 1
                        grid_anti.col = 1
                        txt_1.Text = grid_anti.TextMatrix(1, 1)
                        txt_1.Width = grid_anti.CellWidth
                        txt_1.Height = grid_anti.CellHeight
                        txt_1.Left = grid_anti.CellLeft + grid_anti.Left
                        txt_1.Top = grid_anti.CellTop + grid_anti.Top
                        txt_1.Visible = True
                        grid_anti.row = 1
                        grid_anti.col = 2
                        txt_2.Text = grid_anti.TextMatrix(1, 2)
                        txt_2.Width = grid_anti.CellWidth
                        txt_2.Height = grid_anti.CellHeight
                        txt_2.Left = grid_anti.CellLeft + grid_anti.Left
                        txt_2.Top = grid_anti.CellTop + grid_anti.Top
                        txt_2.Visible = True
                    End If
                Else
                    'only one organisms
                    grid_anti.Cols = 2
                    grid_anti.row = 1
                        grid_anti.col = 1
                        txt_1.Text = grid_anti.TextMatrix(1, 1)
                        txt_1.Width = grid_anti.CellWidth
                        txt_1.Height = grid_anti.CellHeight
                        txt_1.Left = grid_anti.CellLeft + grid_anti.Left
                        txt_1.Top = grid_anti.CellTop + grid_anti.Top
                        txt_1.Visible = True
              End If
             txt_1.SetFocus
             End If
    
    ElseIf mod_flag Then
            
'**************** Commented and added by Shraddha on 30/06/09 *****************
' If ((Len(txt_organ1.Text) >= 1 Or Len(TXT_ORGAN2.Text) >= 1 Or Len(TXT_ORGAN3.Text) >= 1)) Then
If ((Len(cmb_org1.Text) >= 1 Or Len(cmb_org2.Text) >= 1 Or Len(cmb_org3.Text) >= 1)) Then
'******************************************************************************
      
                    grid_anti.Rows = 1
                    grid_anti.Cols = 4
                    grid_anti.FormatString = "    ANTIBIOTIC                               |^         1           |^          2         |^              3      "

      
                    sel_antiqr(0) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
                    sel_antiqr(1) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
                    Set sel_antirs = sel_antiqr.OpenResultset(1, 3)

                    While Not sel_antirs.EOF
                        
                        retantibodies_qr(0) = mod_reqno
                        Dim ANTI As String
                        ANTI = Trim(sel_antirs("antibdesc"))
                        retantibodies_qr(1) = Trim(sel_antirs("antibdesc"))
                        'grid_anti.AddItem sel_antirs("antibdesc")
                        Set RETRESULT_RS = retantibodies_qr.OpenResultset(1, 3)
                        
                        If RETRESULT_RS.RowCount > 0 Then
                            grid_anti.AddItem RETRESULT_RS("antibiotic") & vbTab & RETRESULT_RS("antibio1") & vbTab & RETRESULT_RS("antibio2") & vbTab & RETRESULT_RS("antibio3")
                        Else
                            grid_anti.AddItem sel_antirs("antibdesc")
                        End If
                        sel_antirs.MoveNext
                    Wend
      
      
      
      'With grid_anti
                
                'sel_forupd2_qr(0) = mod_reqno
                'Set RETRESULT_RS = sel_forupd2_qr.OpenResultset(1, 3)
                
                'While Not RETRESULT_RS.EOF
                 '      grid_anti.AddItem RETRESULT_RS("antibiotic") & vbTab & RETRESULT_RS("antibio1") & vbTab & RETRESULT_RS("antibio2") & vbTab & RETRESULT_RS("antibio3")
                 '       RETRESULT_RS.MoveNext
                  '  Wend
      
      '***** Shraddha 30/06/09 *********
      'If Len(txt_organ1.Text) >= 1 Then
      If Len(cmb_org1.Text) >= 1 Then
      
                 'If Len(TXT_ORGAN2.Text) >= 1 Then
                 If Len(cmb_org2.Text) >= 1 Then
                 
                     'If Len(TXT_ORGAN3.Text) >= 1 Then
                     If Len(cmb_org3.Text) >= 1 Then
    '**********************************
                        'three organisms
                        grid_anti.row = 1
                        grid_anti.col = 1
                        txt_1.Text = grid_anti.TextMatrix(1, 1)
                        txt_1.Width = grid_anti.CellWidth
                        txt_1.Height = grid_anti.CellHeight
                        txt_1.Left = grid_anti.CellLeft + grid_anti.Left
                        txt_1.Top = grid_anti.CellTop + grid_anti.Top
                        txt_1.Visible = True
                        grid_anti.row = 1
                        grid_anti.col = 2
                        txt_2.Text = grid_anti.TextMatrix(1, 2)
                        txt_2.Width = grid_anti.CellWidth
                        txt_2.Height = grid_anti.CellHeight
                        txt_2.Left = grid_anti.CellLeft + grid_anti.Left
                        txt_2.Top = grid_anti.CellTop + grid_anti.Top
                        txt_2.Visible = True
                        grid_anti.row = 1
                        grid_anti.col = 3
                        txt_3.Text = grid_anti.TextMatrix(1, 3)
                        txt_3.Width = grid_anti.CellWidth
                        txt_3.Height = grid_anti.CellHeight
                        txt_3.Left = grid_anti.CellLeft + grid_anti.Left
                        txt_3.Top = grid_anti.CellTop + grid_anti.Top
                        txt_3.Visible = True
                    Else
                        'two organisms
                        grid_anti.Cols = 3
                        grid_anti.row = 1
                        grid_anti.col = 1
                        txt_1.Text = grid_anti.TextMatrix(1, 1)
                        txt_1.Width = grid_anti.CellWidth
                        txt_1.Height = grid_anti.CellHeight
                        txt_1.Left = grid_anti.CellLeft + grid_anti.Left
                        txt_1.Top = grid_anti.CellTop + grid_anti.Top
                        txt_1.Visible = True
                        grid_anti.row = 1
                        grid_anti.col = 2
                        txt_2.Text = grid_anti.TextMatrix(1, 2)
                        txt_2.Width = grid_anti.CellWidth
                        txt_2.Height = grid_anti.CellHeight
                        txt_2.Left = grid_anti.CellLeft + grid_anti.Left
                        txt_2.Top = grid_anti.CellTop + grid_anti.Top
                        txt_2.Visible = True
                    End If
                Else
                    'only one organisms
                    grid_anti.Cols = 2
                    grid_anti.row = 1
                        grid_anti.col = 1
                        txt_1.Text = grid_anti.TextMatrix(1, 1)
                        txt_1.Width = grid_anti.CellWidth
                        txt_1.Height = grid_anti.CellHeight
                        txt_1.Left = grid_anti.CellLeft + grid_anti.Left
                        txt_1.Top = grid_anti.CellTop + grid_anti.Top
                        txt_1.Visible = True
              End If
             txt_1.SetFocus
             End If
                        
'                        grid_anti.row = 1
'                        grid_anti.Col = 1
'                        txt_1.Text = grid_anti.TextMatrix(1, 1)
'                        txt_1.Width = grid_anti.CellWidth
'                        txt_1.Height = grid_anti.CellHeight
'                        txt_1.Left = grid_anti.CellLeft + grid_anti.Left
'                        txt_1.Top = grid_anti.CellTop + grid_anti.Top
'                        txt_1.Visible = False
'                        grid_anti.row = 1
'                        grid_anti.Col = 2
'                        txt_2.Text = grid_anti.TextMatrix(1, 2)
'                        txt_2.Width = grid_anti.CellWidth
'                        txt_2.Height = grid_anti.CellHeight
'                        txt_2.Left = grid_anti.CellLeft + grid_anti.Left
'                        txt_2.Top = grid_anti.CellTop + grid_anti.Top
'                        txt_2.Visible = False
'                        grid_anti.row = 1
'                        grid_anti.Col = 3
'                        txt_3.Text = grid_anti.TextMatrix(1, 3)
'                        txt_3.Width = grid_anti.CellWidth
'                        txt_3.Height = grid_anti.CellHeight
'                        txt_3.Left = grid_anti.CellLeft + grid_anti.Left
'                        txt_3.Top = grid_anti.CellTop + grid_anti.Top
'                        txt_3.Visible = False

                        
      
      'End With
      
            
 End If
            
            
            '********************** COMMENTED BY KANCHAN
'                If ((Len(txt_organ1.Text) >= 1 Or Len(TXT_ORGAN2.Text) >= 1 Or Len(TXT_ORGAN3.Text) >= 1)) Then
'               ' And Trim(grid_anti.TextMatrix(1, 1)) = "")
'
'                    sel_antiqr(0) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
'                    sel_antiqr(1) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
'                    Set sel_antirs = sel_antiqr.OpenResultset(1, 3)
'
'                   'three organisms
'                    grid_anti.Rows = 1
'                    grid_anti.Cols = 2
'                    grid_anti.FormatString = "    ANTIBIOTIC                               |^         1           |^          2         |^              3      "
'                    While Not sel_antirs.EOF
'                        grid_anti.AddItem sel_antirs("antibdesc")
'                        sel_antirs.MoveNext
'                    Wend
'                 End If
'            '**********************
'        If grid_anti.Rows > 1 Then
'            'means previous data is available
'            If Len(txt_organ1.Text) >= 1 Then
'                 If Len(TXT_ORGAN2.Text) >= 1 Then
'                     If Len(TXT_ORGAN3.Text) >= 1 Then
'                        'three organisms
'                        grid_anti.row = 1
'                        grid_anti.Col = 1
'                        txt_1.Text = grid_anti.TextMatrix(1, 1)
'                        txt_1.Width = grid_anti.CellWidth
'                        txt_1.Height = grid_anti.CellHeight
'                        txt_1.Left = grid_anti.CellLeft + grid_anti.Left
'                        txt_1.Top = grid_anti.CellTop + grid_anti.Top
'                        txt_1.Visible = True
'                        grid_anti.row = 1
'                        grid_anti.Col = 2
'                        txt_2.Text = grid_anti.TextMatrix(1, 2)
'                        txt_2.Width = grid_anti.CellWidth
'                        txt_2.Height = grid_anti.CellHeight
'                        txt_2.Left = grid_anti.CellLeft + grid_anti.Left
'                        txt_2.Top = grid_anti.CellTop + grid_anti.Top
'                        txt_2.Visible = True
'                        grid_anti.row = 1
'                        grid_anti.Col = 3
'                        txt_3.Text = grid_anti.TextMatrix(1, 3)
'                        txt_3.Width = grid_anti.CellWidth
'                        txt_3.Height = grid_anti.CellHeight
'                        txt_3.Left = grid_anti.CellLeft + grid_anti.Left
'                        txt_3.Top = grid_anti.CellTop + grid_anti.Top
'                        txt_3.Visible = True
'                    Else
'                        'two organisms
'                        grid_anti.Cols = 3
'                        grid_anti.row = 1
'                        grid_anti.Col = 1
'                        txt_1.Text = grid_anti.TextMatrix(1, 1)
'                        txt_1.Width = grid_anti.CellWidth
'                        txt_1.Height = grid_anti.CellHeight
'                        txt_1.Left = grid_anti.CellLeft + grid_anti.Left
'                        txt_1.Top = grid_anti.CellTop + grid_anti.Top
'                        txt_1.Visible = True
'                        grid_anti.row = 1
'                        grid_anti.Col = 2
'                        txt_2.Text = grid_anti.TextMatrix(1, 2)
'                        txt_2.Width = grid_anti.CellWidth
'                        txt_2.Height = grid_anti.CellHeight
'                        txt_2.Left = grid_anti.CellLeft + grid_anti.Left
'                        txt_2.Top = grid_anti.CellTop + grid_anti.Top
'                        txt_2.Visible = True
'                    End If
'                Else
'                    'only one organisms
'                    grid_anti.Cols = 2
'                    grid_anti.row = 1
'                        grid_anti.Col = 1
'                        txt_1.Text = grid_anti.TextMatrix(1, 1)
'                        txt_1.Width = grid_anti.CellWidth
'                        txt_1.Height = grid_anti.CellHeight
'                        txt_1.Left = grid_anti.CellLeft + grid_anti.Left
'                        txt_1.Top = grid_anti.CellTop + grid_anti.Top
'                        txt_1.Visible = True
'              End If
'             txt_1.SetFocus
'             End If
'
'
'        Else
'            ' load new antibiotics
'            If Len(txt_organ1.Text) >= 1 Then
'                sel_antiqr(0) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
'                sel_antiqr(1) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
'                Set sel_antirs = sel_antiqr.OpenResultset(1, 3)
'                If Len(TXT_ORGAN2.Text) >= 1 Then
'                    If Len(TXT_ORGAN3.Text) >= 1 Then
'                        'three organisms
'                        grid_anti.Rows = 1
'                        grid_anti.Cols = 3
'                        grid_anti.FormatString = "    ANTIBIOTIC                               |^         1           |^          2         |^              3      "
'                        While Not sel_antirs.EOF
'                            grid_anti.AddItem sel_antirs("antibdesc")
'                            sel_antirs.MoveNext
'                        Wend
'                        grid_anti.row = 1
'                        grid_anti.Col = 1
'                        txt_1.Width = grid_anti.CellWidth
'                        txt_1.Height = grid_anti.CellHeight
'                        txt_1.Left = grid_anti.CellLeft + grid_anti.Left
'                        txt_1.Top = grid_anti.CellTop + grid_anti.Top
'                        grid_anti.row = 1
'                        grid_anti.Col = 2
'                        txt_2.Width = grid_anti.CellWidth
'                        txt_2.Height = grid_anti.CellHeight
'                        txt_2.Left = grid_anti.CellLeft + grid_anti.Left
'                        txt_2.Top = grid_anti.CellTop + grid_anti.Top
'                        grid_anti.row = 1
'                        grid_anti.Col = 3
'                        txt_3.Width = grid_anti.CellWidth
'                        txt_3.Height = grid_anti.CellHeight
'                        txt_3.Left = grid_anti.CellLeft + grid_anti.Left
'                        txt_3.Top = grid_anti.CellTop + grid_anti.Top
'                        txt_3.Visible = True
'                        txt_2.Visible = True
'                        txt_1.Visible = True
'                    Else
'                        'two organisms
'                        grid_anti.Rows = 1
'                        grid_anti.Cols = 2
'                        grid_anti.FormatString = "    ANTIBIOTIC                               |^         1           |^          2         "
'                        While Not sel_antirs.EOF
'                            grid_anti.AddItem sel_antirs("antibdesc")
'                            sel_antirs.MoveNext
'                        Wend
'                            grid_anti.row = 1
'                            grid_anti.Col = 1
'                            txt_1.Width = grid_anti.CellWidth
'                            txt_1.Height = grid_anti.CellHeight
'                            txt_1.Left = grid_anti.CellLeft + grid_anti.Left
'                            txt_1.Top = grid_anti.CellTop + grid_anti.Top
'                            grid_anti.row = 1
'                            grid_anti.Col = 2
'                            txt_2.Width = grid_anti.CellWidth
'                            txt_2.Height = grid_anti.CellHeight
'                            txt_2.Left = grid_anti.CellLeft + grid_anti.Left
'                            txt_2.Top = grid_anti.CellTop + grid_anti.Top
'                            txt_3.Visible = False
'                            txt_2.Visible = True
'                            txt_1.Visible = True
'                        End If
'                Else
'                'only one organisms
'                    grid_anti.Rows = 1
'                    grid_anti.Cols = 2
'                    grid_anti.FormatString = "    ANTIBIOTIC                               |^         1           "
'                    While Not sel_antirs.EOF
'                        grid_anti.AddItem sel_antirs("antibdesc")
'                        sel_antirs.MoveNext
'                    Wend
'                    grid_anti.row = 1
'                    grid_anti.Col = 1
'                    txt_1.Width = grid_anti.CellWidth
'                    txt_1.Height = grid_anti.CellHeight
'                    txt_1.Left = grid_anti.CellLeft + grid_anti.Left
'                    txt_1.Top = grid_anti.CellTop + grid_anti.Top
'                    txt_1.Visible = True
'                    txt_2.Visible = False
'                    txt_3.Visible = False
'                End If
'
'        Else
'            MsgBox "Atleast 1 organisms is needed", vbOKOnly
'            Exit Sub
'        End If
'        End If
'      '*********************KANCHAN
    End If 'FIRST IF
    Exit Sub
MSG:
        MsgBox ERROR(Err)
End Sub

    Public Sub btnenabl(cap As String)
If cap = "NEW" Then
    If Not USRNAME_RS.EOF Then
        If USRNAME_RS("SAVE") = "YES" Then
            cmd_commit.Enabled = False
            
             cmd_save.Enabled = True
             cmd_save.Visible = True 'MODIFIED BY MRUNAL 23_2
        End If
    End If
    cmd_cancel.Enabled = True
     
    CMD_NEW.Enabled = False
    cmd_modify.Enabled = False
       cmd_print.Enabled = False
    'cmd_commit.Enabled = False
    cmd_view.Enabled = False
    cmd_fullreport.Enabled = True
ElseIf cap = "VIEW" Then
    cmd_cancel.Enabled = True
  
    If Not USRNAME_RS.EOF Then
        If USRNAME_RS("MODIFY") = "YES" Then
            cmd_modify.Enabled = True
            cmd_commit.Enabled = True
            cmd_save.Enabled = False
        Else
            cmd_modify.Enabled = False
            cmd_commit.Enabled = False
              cmd_save.Enabled = True
        End If
    End If
    CMD_NEW.Enabled = False
' Mdified by mrunal on 21_2_13
          
           '------end
    cmd_print.Enabled = True
    cmd_view.Enabled = False
    cmd_fullreport.Enabled = True
ElseIf cap = "SAVE" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_previous.Enabled = True
      cmd_modify.Enabled = False
      cmd_commit.Enabled = False
      
    If Not USRNAME_RS.EOF Then
        If USRNAME_RS("NEW") = "YES" Then
            CMD_NEW.Enabled = True
        Else
            CMD_NEW.Enabled = False
        End If
    End If
    ' Mdified by mrunal on 21_2_13
           cmd_save.Enabled = True
           '------end
    cmd_print.Enabled = True
    If Not USRNAME_RS.EOF Then
        If USRNAME_RS("COMMIT") = "YES" Then
            cmd_commit.Enabled = True
        Else
            cmd_commit.Enabled = False
        End If
    End If
    If Not USRNAME_RS.EOF Then
        If USRNAME_RS("VIEW") = "YES" Then
            cmd_view.Enabled = True
             cmd_save.Enabled = True
        Else
            cmd_view.Enabled = False
             cmd_save.Enabled = False
        End If
    End If
    cmd_fullreport.Enabled = True
    'MSFlexGrid1.Enabled = False 'Shraddha Pandit 29/10/08
ElseIf cap = "INI" Then
    cmd_cancel.Enabled = True
     cmd_back.Enabled = True
    cmd_previous.Enabled = True
    cmd_modify.Enabled = False
    cmd_commit.Enabled = False
      cmd_save.Enabled = False
    If Not USRNAME_RS.EOF Then
        If USRNAME_RS("NEW") = "YES" Then
            CMD_NEW.Enabled = True
        Else
            CMD_NEW.Enabled = False
        End If
    End If
   ' Mdified by mrunal on 21_2_13
         
           '------end
    cmd_print.Enabled = False

    If Not USRNAME_RS.EOF Then
        If USRNAME_RS("VIEW") = "YES" Then
            cmd_view.Enabled = True
        Else
            cmd_view.Enabled = False
        End If
    End If
    cmd_fullreport.Enabled = True
ElseIf cap = "COMMIT" Then
    ' Mdified by mrunal on 21_2_13
         
           '------end
    cmd_cancel.Enabled = False
    cmd_back.Enabled = True
    cmd_previous.Enabled = True
    If Not USRNAME_RS.EOF Then
        If USRNAME_RS("NEW") = "YES" Then
            CMD_NEW.Enabled = True
        Else
            CMD_NEW.Enabled = False
        End If
    End If
    cmd_modify.Enabled = False
    cmd_print.Enabled = True
      If Not USRNAME_RS.EOF Then
        If USRNAME_RS("VIEW") = "YES" Then
            cmd_view.Enabled = True
        Else
            cmd_view.Enabled = False
        End If
    End If
    cmd_fullreport.Enabled = True
ElseIf cap = "print" Then
     ' Mdified by mrunal on 21_2_13
           cmd_save.Enabled = False
           '------end
       cmd_back.Enabled = True
    cmd_previous.Enabled = True
     cmd_modify.Enabled = False
     cmd_print.Enabled = False
     cmd_cancel.Enabled = True
     CMD_NEW.Enabled = False
End If

'If cap = "NEW" Then
'    cmd_save.Enabled = True
'    cmd_cancel.Enabled = True
'    cmd_back.Enabled = True
'    CMD_NEW.Enabled = False
'    cmd_modify.Enabled = False
'    cmd_print.Enabled = False
'    cmd_commit.Enabled = False
'    cmd_view.Enabled = False
'    cmd_fullreport.Enabled = False
'ElseIf cap = "VIEW" Then
'    cmd_cancel.Enabled = True
'    cmd_back.Enabled = True
'    cmd_modify.Enabled = True
'    CMD_NEW.Enabled = False
'    cmd_save.Enabled = False
'    cmd_print.Enabled = True
'    cmd_commit.Enabled = False
'    cmd_view.Enabled = False
'    cmd_fullreport.Enabled = True
'ElseIf cap = "SAVE" Then
'    cmd_cancel.Enabled = True
'    cmd_back.Enabled = True
'    cmd_modify.Enabled = False
'    CMD_NEW.Enabled = True
'    cmd_save.Enabled = False
'    cmd_print.Enabled = True
'    cmd_commit.Enabled = True
'    cmd_view.Enabled = True
'    cmd_fullreport.Enabled = False
'ElseIf cap = "INI" Then
'    cmd_cancel.Enabled = True
'    cmd_back.Enabled = True
'    cmd_modify.Enabled = False
'    CMD_NEW.Enabled = True
'    cmd_save.Enabled = False
'    cmd_print.Enabled = False
'    cmd_commit.Enabled = False
'    cmd_view.Enabled = True
'    cmd_fullreport.Enabled = False
'ElseIf cap = "COMMIT" Then
'    cmd_save.Enabled = False
'    cmd_cancel.Enabled = False
'    cmd_back.Enabled = True
'    CMD_NEW.Enabled = True
'    cmd_modify.Enabled = False
'    cmd_print.Enabled = True
'    cmd_commit.Enabled = False
'    cmd_view.Enabled = True
'    cmd_fullreport.Enabled = False
'ElseIf cap = "print" Then
'    cmd_save.Enabled = False
'    cmd_back.Enabled = True
'    cmd_modify.Enabled = False
'    cmd_print.Enabled = False
'    cmd_cancel.Enabled = True
'    CMD_NEW.Enabled = False
'ElseIf cap = "cancel" Then
'    cmd_save.Enabled = False
'    cmd_back.Enabled = True
'    cmd_modify.Enabled = False
'    cmd_print.Enabled = False
'    cmd_cancel.Enabled = True
'    CMD_NEW.Enabled = False
'End If
End Sub

Private Sub initialize_combo()
On Error Resume Next
    cmb_gnb.ListIndex = 0
    cmb_gnc.ListIndex = 0
    cmb_gpb.ListIndex = 0
    cmb_gpc.ListIndex = 0
    cmb_patho.ListIndex = 0
End Sub

Private Sub clearupperpanel()
'upper panel
    lblicu.Visible = False
    MaskEdBox1.Text = ""
    cmb_reqno.clear
    cmb_modality.clear
    txt_deptcode.Text = ""
    cmb_dept.clear
    txt_microno.Text = ""
    txt_name.Text = ""
    
'    '***** Shraddha 22/03/11 *****
'    msk_coll_dt.Text = "__/__/____"
'    msk_coll_tm.Text = "__:__"
'    '*****************************
    
    txt_age.Text = ""
    txt_spm.Text = ""
    txt_sex.Text = ""
    txt_category.Text = ""
    txt_status.Text = ""
    txt_ward.Text = ""
    txt_bedno.Text = ""
    txt_pridoc.Text = ""
    txt_modcode.Text = ""
    txt_refdoc.Text = ""
    CMB_PROC.clear
    cmb_enteryby.Text = ""
    cmb_final.Text = ""
'   grid_anti.Row = 1
    txt_ziehl.Text = ""
    cmb_india.ListIndex = -1
End Sub
'for procdure bacteriology
Private Sub clearlowerpanel1()
    lblicu.Visible = False
    txt_epith.Text = ""
    txt_pus.Text = ""
    txt_other.Text = ""
    'txt_ziehl.Text = ""
    txt_koh.Text = ""
    
    '***** Shraddha 30/06/09 ******
    'txt_organ1.Text = ""
    'TXT_ORGAN2.Text = ""
    'TXT_ORGAN3.Text = ""
    
    cmb_org1.Text = ""
    cmb_org2.Text = ""
    cmb_org3.Text = ""
    '******************************
    txt_1.Text = ""
    txt_2.Text = ""
    txt_3.Text = ""
    txt_report.Text = ""
    txt_comment.Text = ""
    initialize_combo
    grid_anti.clear
    grid_dmbanti.clear
    cmb_final.Text = ""
    cmb_enteryby.Text = ""
    txt_ziehl.Text = ""
    cmb_india.ListIndex = -1
End Sub
'for procedure clinical microbiology

Private Sub clearlowerpanel2()
On Error GoTo MSG
    lblicu.Visible = False
    txt_urinecolour.Text = ""
    txt_gravity.Text = ""
    
    'txt_ph.Text = ""
    'txt_glucose.Text = ""
    'txt_ketones.Text = ""
    'txt_bilirubin.Text = ""
    'txt_protins.Text = ""
    'txt_uro.Text = ""
    'txt_blood.Text = ""
    
    txt_casts.Text = ""
    txt_urinother.Text = ""
    txt_urinecomment.Text = ""
    txt_bacteria.Text = ""
    txt_wbc.Text = ""
    
    
    cmb_blood.Text = ""
    cmb_glucose.Text = ""
    cmb_ph.Text = ""
    cmb_ketones.Text = ""
    cmb_proteins.Text = ""
    cmb_uroblngn.Text = ""
    cmb_nitrite.Text = "" 'Shraddha 21/03/11
    
    cmb_bilirubin.Text = ""
    
    txt_rbc.Text = ""
    txt_squamous.Text = ""
    txt_stoolcolour.Text = ""
    txt_consistency.Text = ""
    txt_puscells.Text = ""
    txt_stoolrbc.Text = ""
    txt_stoolother.Text = ""
    txt_stoolcomments.Text = ""
    txt_parasties.Text = ""
    txt_larvae.Text = ""
    txt_starch.Text = ""
    txt_occult.Text = ""
    txt_fat.Text = ""
    txt_mucus.Text = ""
    txt_Ova.Text = ""
    txt_sblood.Text = ""
    txt_cysts.Text = ""
    txt_microphages.Text = ""
    txt_ziehl.Text = ""
    cmb_india.ListIndex = -1
    'cmb_final.Text = ""
    'cmb_enteryby.Text = ""
    Exit Sub
MSG:
    MsgBox ERROR(Err)
End Sub
'for other procedure
Private Sub clearlowerpanel3()
On Error GoTo MSG
   lblicu.Visible = False
   ' txt_refer_service.Text = ""
    txt_test_perform.Text = ""
    txt_result.Text = ""
    cmb_final.Text = ""
    cmb_enteryby.Text = ""
    txt_ziehl.Text = ""
    cmb_india.ListIndex = -1
    Exit Sub
MSG:
    MsgBox ERROR(Err)
End Sub
Private Sub clearlowerpanel4()
On Error GoTo MSG
    lblicu.Visible = False
    txt_c_report.Text = ""
    txt_c_comment.Text = ""
    'add on 16/05
    cmb_stain.Text = ""
    cmb_organisms.Text = ""
    cmb_organisms.clear
    cmb_final.Text = ""
    cmb_enteryby.Text = ""
    txt_ziehl.Text = ""
    cmb_india.ListIndex = -1
    txt_addendum.Text = ""
    cmb_addendumby.Text = ""
    Exit Sub
MSG:
    MsgBox ERROR(Err)
End Sub
Private Sub clearlowerpanel5()
    txtSerTestPerform.Text = ""
    cmbSerResult.clear
    cmbSerRange.clear
    cmbProcalci.clear
    cmbProCalCon.clear
    txtInterpretation.Text = ""
    txtLimitation.Text = ""
    txtMethod.Text = ""
    txtInBold.Text = ""
End Sub
'common procedure for new and view mode
Private Sub getupperpanel()
'On Error GoTo MSG
        'THIS TO GET DEPTCODE AND MODALITYCODE
        sel_micro_qr(0) = mod_reqno
        Set sel_micro_rs = sel_micro_qr.OpenResultset(1, 3)
        While Not sel_micro_rs.EOF
            txt_microno.Text = Trim(sel_micro_rs("pathno"))
            sel_micro_rs.MoveNext
        Wend
        S1 = Mid(cmb_reqno.Text, 1, 3) 'TO GET DEPTCODE
        S2 = Mid(cmb_reqno.Text, 5, 2) ' TO GET MODALITY CODE
        
        SEL_MODALITY_QR(0) = S1
        SEL_MODALITY_QR(1) = S2
        Set SEL_MODALITY_RS = SEL_MODALITY_QR.OpenResultset(1, 3)
        While Not SEL_MODALITY_RS.EOF
            txt_modcode.Text = Trim(SEL_MODALITY_RS("MODLITYDES"))
                    cmb_modality.Text = Trim(SEL_MODALITY_RS("MODLITYDES")) & "$" & Mid(cmb_reqno.Text, 5, 2)
            SEL_MODALITY_RS.MoveNext
        Wend
        txt_deptcode.Text = S1
        txt_modcode.Text = S2

        If Opt_patient.Value Then
            SEL_detail_qr(0) = MaskEdBox1.FormattedText
            Set SEL_detail_rs = SEL_detail_qr.OpenResultset(1, 3)
            
            
            Dim rn
            rn = Trim(cmb_reqno.Text)
            Dim S As Integer
            S = InStr(1, cmb_reqno, "$")
            rn = req_chglen(Mid(Trim(cmb_reqno), 1, S - 1))
            
            
            SELSPONANTIBQR(0) = rn
            Set SELSPONANTIBRS = SELSPONANTIBQR.OpenResultset(1, 3)
            
            If SEL_detail_rs.RowCount = 0 Then
                SEL_detail_rs.Close
                Exit Sub
            Else
                txt_sex.Text = SEL_detail_rs("sex")
                txt_name.Text = Trim(SEL_detail_rs("first_name")) & " " & Trim(SEL_detail_rs("mid_name")) & " " & Trim(SEL_detail_rs("last_name"))
                txt_age = DateDiff("YYYY", SEL_detail_rs("BIRTH_DATE"), CStr(DATE11))
                txt_pridoc.Text = SEL_detail_rs("ud_name")
                txt_category.Text = SEL_detail_rs("pat_catg")
                If Report_microbiology1.Text2.Text <> "" Then
                Else
                   txt_spm.Text = SELSPONANTIBRS("specimen")
                End If
             
                SEL_detail_rs.Close
            End If
        End If
            REFDR_QR(0) = MaskEdBox1.FormattedText
          ' If Report_microbiology.Text2.Text = "1" Then
            ' REFDR_QR(1) = Mid(Trim(cmb_reqno), 1, S - 1)
           ' Else
            REFDR_QR(1) = mod_reqno
           ' End If
            Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
        If REFDR_RS.RowCount > 0 Then
        'put check here also 20/12/02
        '***********************************
        
            txt_refdoc.Text = REFDR_RS("REFBYDOC")
            If Not callstatus Then
                If Not new_flag = True Then 'Shraddha 17/11/08
                    If REFDR_RS("COMMITFLAG") = "Y" And REFDR_RS("REQNSTAT") = "F" Then
                        MsgBox "Final Report is Generated, You cannot Modify the Final Report"
                        cmd_modify.Enabled = False
                        cmd_save.Enabled = False 'MODIFIED BY MRUNAL ON  23_2
                        cmd_commit.Enabled = False  'MODIFIED BY MRUNAL ON  23_2
                        txt_addendum.Enabled = True
                        cmb_addendumby.Enabled = True
                        'ADDENDUMFALG = True
                    Else
                        cmd_modify.Enabled = True
                         cmd_commit.Enabled = True
                    End If
                End If
            End If
        End If
            sel_bed_qr(0) = MaskEdBox1.FormattedText
            Set sel_bed_rs = sel_bed_qr.OpenResultset(1, 3)
            
        If REFDR_RS("PATNSTAT") = "IP" Then
            
            txt_ward.Visible = True
            txt_bedno.Visible = True
            lbl_ward.Visible = True
            
            sel_micro_rs.MoveFirst
            txt_ward.Text = sel_micro_rs("WARD_NAME")
            txt_bedno.Text = sel_micro_rs("bedno")
            txt_status.Text = REFDR_RS("PATNSTAT")
            sel_bed_rs.Close
        Else
            txt_status.Text = REFDR_RS("PATNSTAT")
            txt_ward.Visible = False
            txt_bedno.Visible = False
            lbl_ward.Visible = False
          
            sel_bed_rs.Close
        End If
        CMB_PROC.clear
        'for procedure previously selected
        
        sel_proccode_qr(0) = mod_reqno
        Set sel_proccode_rs = sel_proccode_qr.OpenResultset(1, 3)
        If Not sel_proccode_rs.EOF Then
            While Not sel_proccode_rs.EOF
                CMB_PROC.AddItem Trim(sel_proccode_rs(1))
                sel_proccode_rs.MoveNext
            Wend
            CMB_PROC.ListIndex = 0
            sel_proccode_rs.Close
'        Else
'            MsgBox "No Procedure Selected Previously", vbOKOnly
'            sel_proccode_rs.Close
'            Exit Sub
        End If
        Exit Sub
MSG:
        MsgBox ERROR(Err)
End Sub

'****************** Shraddha Pandit 16/06/2009 *********************
'Procedure for Previous reqno
Private Sub getupperpanel_prev()
'On Error GoTo MSG
        'THIS TO GET DEPTCODE AND MODALITYCODE
        
        Dim sel_rreq_qr As rdoQuery
        Dim sel_rreq_rs As rdoResultset
        
        Dim G As Integer
        Dim req, sqlstr As String
                
        sqlstr = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
        Set sel_rreq_qr = rdb.CreateQuery("", sqlstr)
        
    
        G = InStr(1, cmb_reqno, "$")
        req = req_chglen(Mid(Trim(cmb_reqno), 1, G - 1))
        sel_micro_qr(0) = req
        Set sel_micro_rs = sel_micro_qr.OpenResultset(1, 3)
        While Not sel_micro_rs.EOF
            txt_microno.Text = Trim(sel_micro_rs("pathno"))
            sel_micro_rs.MoveNext
        Wend
        S1 = Mid(cmb_reqno.Text, 1, 3) 'TO GET DEPTCODE
        S2 = Mid(cmb_reqno.Text, 5, 2) ' TO GET MODALITY CODE
        
        SEL_MODALITY_QR(0) = S1
        SEL_MODALITY_QR(1) = S2
        Set SEL_MODALITY_RS = SEL_MODALITY_QR.OpenResultset(1, 3)
        While Not SEL_MODALITY_RS.EOF
            txt_modcode.Text = Trim(SEL_MODALITY_RS("MODLITYDES"))
            SEL_MODALITY_RS.MoveNext
        Wend
        txt_deptcode.Text = S1
        txt_modcode.Text = S2
                
        sel_rreq_qr(0) = req
        Set sel_rreq_rs = sel_rreq_qr.OpenResultset(1, 3)
        If Len(Trim(sel_rreq_rs("caseno"))) = 8 Then 'Patient
               
            MaskEdBox1.Text = Trim(sel_rreq_rs("caseno"))
            sel_moadlityqr1(0) = Trim(sel_rreq_rs("modalitycd"))
            Set sel_modalityrs1 = sel_moadlityqr1.OpenResultset(1, 3)
            If sel_modalityrs1.RowCount > 0 Then
                cmb_modality.clear
                'While Not sel_modalityrs1.EOF
                    cmb_modality.AddItem Trim(sel_modalityrs1("modlitydes")) & "$" & Trim(sel_modalityrs1("modalitycd"))
                    cmb_modality.Text = Trim(sel_modalityrs1("modlitydes")) & "$" & Trim(sel_modalityrs1("modalitycd"))
                    'sel_modalityrs1.MoveNext
                'Wend
                sel_modalityrs1.Close
            Else
                sel_modalityrs1.Close
            End If
            
            
            
           ' cmb_modality.Text = Trim(sel_rreq_rs("modalitycd"))
            SEL_detail_qr(0) = Trim(sel_rreq_rs("caseno"))
            Set SEL_detail_rs = SEL_detail_qr.OpenResultset(1, 3)
            
            Dim rn
            'rn = Trim(cmb_reqno.Text)
            'Dim s As Integer
            's = InStr(1, cmb_reqno, "$")
            'rn = req_chglen(Mid(Trim(cmb_reqno), 1, s - 1))
            rn = req
            
            SELSPONANTIBQR(0) = rn
            Set SELSPONANTIBRS = SELSPONANTIBQR.OpenResultset(1, 3)
            
            If SEL_detail_rs.RowCount = 0 Then
                SEL_detail_rs.Close
                Exit Sub
            Else
                txt_sex.Text = SEL_detail_rs("sex")
                txt_name.Text = Trim(SEL_detail_rs("first_name")) & " " & Trim(SEL_detail_rs("mid_name")) & " " & Trim(SEL_detail_rs("last_name"))
                txt_age = DateDiff("YYYY", SEL_detail_rs("BIRTH_DATE"), CStr(DATE11))
                txt_pridoc.Text = SEL_detail_rs("ud_name")
                txt_category.Text = SEL_detail_rs("pat_catg")
                txt_spm.Text = SELSPONANTIBRS("specimen")
                SEL_detail_rs.Close
            End If
        End If
            REFDR_QR(0) = Trim(sel_rreq_rs("caseno"))
                          REFDR_QR(1) = mod_reqno
           
          
            Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
        If REFDR_RS.RowCount > 0 Then
        'put check here also 20/12/02
        '***********************************
        
            txt_refdoc.Text = REFDR_RS("REFBYDOC")
            If Not callstatus Then
                If Not new_flag = True Then 'Shraddha 17/11/08
                    If REFDR_RS("COMMITFLAG") = "Y" Then
                        MsgBox "Final Report is Generated, You cannot Modify the Final Report"
                        cmd_modify.Enabled = False
                        cmd_commit.Enabled = False
                        txt_addendum.Enabled = True
                        cmb_addendumby.Enabled = True
                    ' Mdified by mrunal on 21_2_13
           cmd_save.Enabled = True
           '------end
                        
                        'ADDENDUMFALG = True
                    Else
                        cmd_modify.Enabled = True
                        cmd_commit.Enabled = True
                    End If
                End If
            End If
        End If
            sel_bed_qr(0) = Trim(sel_rreq_rs("caseno"))
            Set sel_bed_rs = sel_bed_qr.OpenResultset(1, 3)
            
        REFDR_QR(0) = Trim(sel_rreq_rs("caseno"))
        REFDR_QR(1) = req
        Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
    
        If REFDR_RS("PATNSTAT") = "IP" Then
            
            txt_ward.Visible = True
            txt_bedno.Visible = True
            lbl_ward.Visible = True
            
            sel_micro_rs.MoveFirst
            txt_ward.Text = sel_micro_rs("WARD_NAME")
            txt_bedno.Text = sel_micro_rs("bedno")
            txt_status.Text = REFDR_RS("PATNSTAT")
            sel_bed_rs.Close
        Else
            txt_status.Text = REFDR_RS("PATNSTAT")
            txt_ward.Visible = False
            txt_bedno.Visible = False
            lbl_ward.Visible = False
      
            sel_bed_rs.Close
        End If
        CMB_PROC.clear
        'for procedure previously selected
        
        sel_proccode_qr(0) = req
        Set sel_proccode_rs = sel_proccode_qr.OpenResultset(1, 3)
        If Not sel_proccode_rs.EOF Then
            While Not sel_proccode_rs.EOF
                CMB_PROC.AddItem Trim(sel_proccode_rs(1))
                sel_proccode_rs.MoveNext
            Wend
            CMB_PROC.ListIndex = 0
            sel_proccode_rs.Close
'        Else
'            MsgBox "No Procedure Selected Previously", vbOKOnly
'            sel_proccode_rs.Close
'            Exit Sub
        End If
        Exit Sub
MSG:
        MsgBox ERROR(Err)
End Sub
'*******************************************************************

'for bacteriology get result from database
Private Sub getresult_bacteriology()
'On Error GoTo MSG
sel_forupd1_qr(0) = mod_reqno
Set sel_forupd1_rs = sel_forupd1_qr.OpenResultset(1, 3)
While Not sel_forupd1_rs.EOF
    txt_epith.Text = Trim(sel_forupd1_rs("epithecell"))
    txt_pus.Text = Trim(sel_forupd1_rs("puscell"))
    txt_other.Text = Trim(sel_forupd1_rs("ofinding"))
    txt_ziehl.Text = Trim(sel_forupd1_rs("ZNSTAIN"))
    For I = 0 To (cmb_gpc.ListCount - 1)
        If Trim(cmb_gpc.List(I)) = Trim(sel_forupd1_rs("gpc")) Then
                cmb_gpc.ListIndex = I
        End If
    Next I
    For I = 0 To (cmb_gnc.ListCount - 1)
        If Trim(cmb_gnc.List(I)) = Trim(sel_forupd1_rs("gnc")) Then
                cmb_gnc.ListIndex = I
        End If
    Next I
    For I = 0 To (cmb_gnb.ListCount - 1)
        If Trim(cmb_gnb.List(I)) = Trim(sel_forupd1_rs("gnb")) Then
                cmb_gnb.ListIndex = I
        End If
    Next I
    For I = 0 To (cmb_gpb.ListCount - 1)
        If Trim(cmb_gpb.List(I)) = Trim(sel_forupd1_rs("gpb")) Then
                cmb_gpb.ListIndex = I
        End If
    Next I
    For I = 0 To (cmb_india.ListCount - 1)
        If Trim(cmb_india.List(I)) = Trim(sel_forupd1_rs("iipc")) Then
                cmb_india.ListIndex = I
        End If
    Next I
    txt_koh.Text = Trim(sel_forupd1_rs("kohamt"))
    '***** Commented and added by Shraddha on 30/06/09 *****
    'txt_organ1.Text = "" & Trim(sel_forupd1_rs("orgso1"))
    'TXT_ORGAN2.Text = "" & Trim(sel_forupd1_rs("orgso2"))
    'TXT_ORGAN3.Text = "" & Trim(sel_forupd1_rs("orgso3"))
    
    cmb_org1.Text = "" & Trim(sel_forupd1_rs("orgso1"))
    cmb_org2.Text = "" & Trim(sel_forupd1_rs("orgso2"))
    cmb_org3.Text = "" & Trim(sel_forupd1_rs("orgso3"))
    '******************************************************
    If sel_forupd1_rs.EOF = False Then
        txt_comment.Text = "" & Trim(sel_forupd1_rs("comments"))
    End If
    If Len(Trim(sel_forupd1_rs("REPORT"))) > 1 Then
        cmb_patho.clear
        cmb_patho.AddItem "Yes"
        cmb_patho.AddItem "No"
        cmb_patho.Text = "No"
        cmb_patho_click
        txt_report.Text = "" & sel_forupd1_rs("REPORT")
        Exit Sub
        
    Else
    'sel_forupd1_rs.Close
    'for loading antibiotics with values
        cmb_patho.clear
        cmb_patho.AddItem "Yes"
        cmb_patho.AddItem "No"
        cmb_patho.Text = "Yes"
        cmb_patho_click
        txt_report.Text = ""
        sel_forupd2_qr(0) = mod_reqno
        Set sel_forupd2_rs = sel_forupd2_qr.OpenResultset(1, 3)
        While Not sel_forupd2_rs.EOF
        grid_anti.clear
        'grid_dmbanti.clear
        
            If Len(Trim(sel_forupd1_rs("orgso1"))) > 0 Then
                If Len(Trim(sel_forupd1_rs("orgso2"))) > 0 Then
                    If Len(Trim(sel_forupd1_rs("orgso3"))) > 0 Then
                       'three organisms
                        grid_dmbanti.Cols = 4
                        grid_dmbanti.Rows = 1
                        grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           |^          2         |^              3      "
                        While Not sel_forupd2_rs.EOF
                            grid_dmbanti.AddItem sel_forupd2_rs("antibiotic") & vbTab & sel_forupd2_rs("antibio1") & vbTab & sel_forupd2_rs("antibio2") & vbTab & sel_forupd2_rs("antibio3")
                            sel_forupd2_rs.MoveNext
                        Wend
                    Else
                    'two organisms
                    grid_dmbanti.Cols = 3
                    grid_dmbanti.Rows = 1
                    grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           |^          2         "
                    While Not sel_forupd2_rs.EOF
                            grid_dmbanti.AddItem sel_forupd2_rs("antibiotic") & vbTab & sel_forupd2_rs("antibio1") & vbTab & sel_forupd2_rs("antibio2")
                            sel_forupd2_rs.MoveNext
                    Wend
                    End If
                Else
                'only one organisms
'                    If MODE = "MODIFY" Then
'                         While Not sel_forupd2_rs.EOF
'                                grid_dmbanti.AddItem sel_forupd2_rs("antibiotic") & vbTab & sel_forupd2_rs("antibio1")
'                                sel_forupd2_rs.MoveNext
'                            Wend
'                    Else
                        grid_dmbanti.Cols = 2
                        grid_dmbanti.Rows = 1
                        grid_dmbanti.FormatString = "    ANTIBIOTIC                               |^         1           "
                            While Not sel_forupd2_rs.EOF
                                grid_dmbanti.AddItem sel_forupd2_rs("antibiotic") & vbTab & sel_forupd2_rs("antibio1")
                                sel_forupd2_rs.MoveNext
                            Wend
'                     End If
                End If
            Else
                sel_forupd2_rs.Close
                Exit Sub
            End If
          
    '     sel_forupd2_rs.MoveNext
         Wend
         sel_forupd2_rs.Close
    End If
    sel_forupd1_rs.MoveNext
Wend
sel_forupd1_rs.Close

'remoanti
    Exit Sub
MSG:
    MsgBox ERROR(Err)
End Sub
'to get result from dmure table for view of urine report
Private Sub getresult_urine()
On Error GoTo MSG
    sel_forupd3_qr(0) = mod_reqno
    Set sel_forupd3_rs = sel_forupd3_qr.OpenResultset(1, 3)
    While Not sel_forupd3_rs.EOF
        txt_urinecolour.Text = Trim(sel_forupd3_rs("color") & "")
        txt_gravity.Text = Trim(sel_forupd3_rs("sp") & "")
        
        'txt_ph.Text = Trim(sel_forupd3_rs("ph") & "")
        cmb_ph.Text = Trim(sel_forupd3_rs("ph") & "")
        'txt_protins.Text = Trim(sel_forupd3_rs("protein") & "")
        cmb_proteins.Text = Trim(sel_forupd3_rs("protein") & "")
        'txt_glucose.Text = Trim(sel_forupd3_rs("glucose") & "")
        cmb_glucose.Text = Trim(sel_forupd3_rs("glucose") & "")
        'txt_ketones.Text = Trim(sel_forupd3_rs("ketones") & "")
        cmb_ketones.Text = Trim(sel_forupd3_rs("ketones") & "")
        'txt_bilirubin.Text = Trim(sel_forupd3_rs("bulirubin") & "")
        cmb_bilirubin.Text = Trim(sel_forupd3_rs("bulirubin") & "")
        'txt_uro.Text = Trim(sel_forupd3_rs("urobilinog") & "")
        cmb_uroblngn.Text = Trim(sel_forupd3_rs("urobilinog") & "")
        
        cmb_nitrite.Text = Trim(sel_forupd3_rs("NITRITE") & "") 'Shraddha 21/03/11
        
        'txt_blood.Text = Trim(sel_forupd3_rs("blood") & "")
        cmb_blood.Text = Trim(sel_forupd3_rs("blood") & "")
        'txt_MicroExam.Text = Trim(sel_forupd3_rs("micexam") & "")
        txt_wbc.Text = Trim(sel_forupd3_rs("wbc") & "")
        txt_rbc.Text = Trim(sel_forupd3_rs("rbc") & "")
        txt_squamous.Text = Trim(sel_forupd3_rs("sqepcell") & "")
        txt_bacteria.Text = Trim(sel_forupd3_rs("bacteria") & "")
        txt_casts.Text = Trim(sel_forupd3_rs("casts") & "")
        txt_urinother.Text = Trim(sel_forupd3_rs("others") & "")
        txt_urinecomment.Text = Trim(sel_forupd3_rs("comments") & "")
        sel_forupd3_rs.MoveNext
    Wend
    sel_forupd3_rs.Close
    Exit Sub
MSG:
    MsgBox ERROR(Err)
End Sub

Private Sub getresult_stool()
On Error GoTo MSG
    sel_forupd4_qr(0) = mod_reqno
    Set sel_forupd4_rs = sel_forupd4_qr.OpenResultset(1, 3)
    While Not sel_forupd4_rs.EOF
        'HERE ALL STOOL FIELD
        txt_stoolcolour.Text = Trim(sel_forupd4_rs("COLOR"))
        txt_consistency.Text = Trim(sel_forupd4_rs("CONSISTENC"))
        txt_mucus.Text = Trim(sel_forupd4_rs("MUCUS"))
        txt_sblood.Text = Trim(sel_forupd4_rs("BLOOD"))
        txt_cysts.Text = Trim(sel_forupd4_rs("CYSTS"))
        txt_microphages.Text = Trim(sel_forupd4_rs("MACOPHAGES"))
        txt_stoolrbc.Text = Trim(sel_forupd4_rs("RBC"))
        txt_Ova.Text = Trim(sel_forupd4_rs("OVA"))
        txt_occult.Text = Trim(sel_forupd4_rs("OCCUTTBLOD"))
        txt_puscells.Text = Trim(sel_forupd4_rs("PUSCELL"))
        txt_fat.Text = Trim(sel_forupd4_rs("FAT"))
        txt_starch.Text = Trim(sel_forupd4_rs("STARCH"))
        txt_larvae.Text = Trim(sel_forupd4_rs("LARVAE"))
        txt_parasties.Text = Trim(sel_forupd4_rs("PARASTIES"))
        txt_stoolcomments.Text = Trim(sel_forupd4_rs("COMMENTS"))
        txt_stoolother.Text = Trim(sel_forupd4_rs("OTHERS"))
        sel_forupd4_rs.MoveNext
    Wend
    sel_forupd4_rs.Close
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub getresult_other()
On Error GoTo MSG
    sel_forupd5_qr(0) = mod_reqno
    Set sel_forupd5_rs = sel_forupd5_qr.OpenResultset(1, 3)
    While Not sel_forupd5_rs.EOF
        'txt_refer_service.Text = sel_forupd5_rs("refserv")
        txt_test_perform.Text = sel_forupd5_rs("testperfor")
        txt_result.Text = sel_forupd5_rs("result")
        
        Dim GBS_QR As rdoQuery
        Dim GBS_RS As rdoResultset
        str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO = ? "
        Set GBS_QR = rdb.CreateQuery("", str)
        GBS_QR(0) = RQN
        Set GBS_RS = GBS_QR.OpenResultset(1, 3)
    
        If Not GBS_RS.EOF Then
            If GBS_RS("PRO_CODE") = "G204" Or GBS_RS("PRO_CODE") = "G262" Then
                   txt_btest.Text = sel_forupd5_rs("testperfor")
                      txt_bresult.Text = sel_forupd5_rs("result")
            End If
            End If
            GBS_RS.Close
        sel_forupd5_rs.MoveNext
    Wend
    Exit Sub
MSG:
    MsgBox ERROR(Err)
End Sub
Private Sub getresult_myco()
'On Error GoTo MSG
    sel_forupd6_qr(0) = mod_reqno
    Set sel_forupd6_rs = sel_forupd6_qr.OpenResultset(1, 3)
    While Not sel_forupd6_rs.EOF
        'cmb_stain.AddItem (Trim(sel_forupd6_rs("stain")))
        cmb_stain.Text = (Trim(sel_forupd6_rs("stain")))
        cmb_organisms.AddItem (Trim(sel_forupd6_rs("organism")))
        cmb_organisms.Text = (Trim(sel_forupd6_rs("organism")))
        txt_c_report.Text = Trim(sel_forupd6_rs("report"))
        txt_c_comment.Text = Trim(sel_forupd6_rs("comment"))
        sel_forupd6_rs.MoveNext
    Wend
    sel_forupd6_rs.Close
    'load antibiotics
    sel_forupd2_qr(0) = mod_reqno
    Set sel_forupd2_rs = sel_forupd2_qr.OpenResultset(1, 3)
    grid_c_anti.clear
    grid_c_anti.Rows = 1
    grid_c_anti.FormatString = "   ANTIMICROBIAL                               |^         1                      "
   
    While Not sel_forupd2_rs.EOF
        grid_c_anti.AddItem sel_forupd2_rs("antibiotic") & vbTab & sel_forupd2_rs("antibio1")
        sel_forupd2_rs.MoveNext
    Wend
    cmb_antiboitic_mb.Visible = False
    txt_c_1.Visible = False
    On Error Resume Next
    grid_c_anti.col = 1
'    '***** Shraddha 12/05/11 *****
'    If mod_flag = True Then
'        grid_c_anti.Enabled = True
'        cmb_antiboitic_mb.Move grid_c_anti.Left + grid_c_anti.CellLeft, grid_c_anti.Top + grid_c_anti.CellTop, grid_c_anti.CellWidth
'        txt_c_1.Move grid_c_anti.Left + grid_c_anti.CellLeft, grid_c_anti.Top + grid_c_anti.CellTop, grid_c_anti.CellWidth
'    End If
'    '********* 12/05/2011 ********
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub set_flag(flag As String) 'used to flag status
If Trim(flag) = "NEW" Then
    new_flag = True
    VIEW_FLAG = False
    mod_flag = False
ElseIf Trim(flag) = "VIEW" Then
    new_flag = False
    VIEW_FLAG = True
    mod_flag = False
ElseIf Trim(flag) = "MOD" Then
    new_flag = False
    VIEW_FLAG = False
    mod_flag = True
Else
    new_flag = False
    VIEW_FLAG = False
    mod_flag = False
End If
End Sub

Private Sub set_position() 'set top position for all three textboxes
On Error GoTo MSG
'    grid_anti.Row = 1
'    grid_anti.Col = 1
'    txt_1.Width = grid_anti.CellWidth
'    txt_1.Height = grid_anti.CellHeight
'    txt_1.Left = grid_anti.CellLeft + grid_anti.Left
'    txt_1.Top = grid_anti.CellTop + grid_anti.Top
'    txt_1.Visible = False
'
'    grid_anti.Row = 1
'    grid_anti.Col = 2
'    txt_2.Width = grid_anti.CellWidth
'    txt_2.Height = grid_anti.CellHeight
'    txt_2.Left = grid_anti.CellLeft + grid_anti.Left
'    txt_2.Top = grid_anti.CellTop + grid_anti.Top
'    txt_2.Visible = False
'
'    grid_anti.Row = 1
'    grid_anti.Col = 3
'    txt_3.Width = grid_anti.CellWidth
'    txt_3.Height = grid_anti.CellHeight
'    txt_3.Left = grid_anti.CellLeft + grid_anti.Left
'    txt_3.Top = grid_anti.CellTop + grid_anti.Top
'    txt_3.Visible = False
'    Exit Sub
MSG:
    MsgBox ERROR(Err)
End Sub
    
Private Sub enable_field()
On Error GoTo MSG
'called when modifying records so disable upper panel
    cmb_dept.Enabled = False
    cmb_modality.Enabled = False
    cmb_dept.Enabled = False
    cmb_reqno.Enabled = False
    MaskEdBox1.Enabled = False
    cmb_enteryby.Enabled = True
    cmb_final.Enabled = True
    MSFlexGrid1.Enabled = True 'Shraddha Pandit 29/10/08
    txt_test.Enabled = True
    txt_MCcomment.Enabled = True
    txt_testresult.Enabled = True
    txt_btest.Enabled = True
    txt_bresult.Enabled = True
'for panel1
    txt_epith.Enabled = True
    txt_pus.Enabled = True
    txt_other.Enabled = True
    txt_koh.Enabled = True
    '***** Shraddha 30/06/09 *****
    'txt_organ1.Enabled = True
    'TXT_ORGAN2.Enabled = True
    'TXT_ORGAN3.Enabled = True
    
    cmb_org1.Enabled = True
    cmb_org2.Enabled = True
    cmb_org3.Enabled = True
    '*****************************
    txt_1.Enabled = True
    txt_2.Enabled = True
    txt_3.Enabled = True
    txt_comment.Enabled = True
    txt_report.Enabled = True
'for panel2
    txt_urinecolour.Enabled = True
    txt_gravity.Enabled = True
    txt_ph.Enabled = True
    txt_glucose.Enabled = True
    txt_blood.Enabled = True
    txt_casts.Enabled = True
    txt_urinother.Enabled = True
    txt_urinecomment.Enabled = True
    txt_bacteria.Enabled = True
    txt_wbc.Enabled = True
    txt_ketones.Enabled = True
    txt_bilirubin.Enabled = True
    txt_rbc.Enabled = True
    txt_protins.Enabled = True
    txt_uro.Enabled = True
    
    cmb_blood.Enabled = True
    cmb_glucose.Enabled = True
    cmb_ph.Enabled = True
    cmb_ketones.Enabled = True
    cmb_proteins.Enabled = True
    cmb_uroblngn.Enabled = True
    cmb_nitrite.Enabled = True 'Shraddha 21/03/11
    
    cmb_bilirubin.Enabled = True
    
    txt_squamous.Enabled = True
    txt_stoolcolour.Enabled = True
    txt_consistency.Enabled = True
    txt_puscells.Enabled = True
    txt_stoolrbc.Enabled = True
    txt_stoolother.Enabled = True
    txt_stoolcomments.Enabled = True
    txt_mucus.Enabled = True
    txt_sblood.Enabled = True
    txt_cysts.Enabled = True
    txt_microphages.Enabled = True
    txt_stoolrbc.Enabled = True
    txt_Ova.Enabled = True
    txt_occult.Enabled = True
    txt_puscells.Enabled = True
    txt_fat.Enabled = True
    txt_starch.Enabled = True
    txt_larvae.Enabled = True
    txt_parasties.Enabled = True
'for panel3
    txt_test_perform.Enabled = True
    txt_result.Enabled = True
'for all combo
    cmb_gnb.Enabled = True
    cmb_gnc.Enabled = True
    cmb_gpb.Enabled = True
    cmb_gpc.Enabled = True
    cmb_india.Enabled = True
    txt_ziehl.Enabled = True
    cmb_patho.Enabled = True
    
  
    'for panel4
    cmb_stain.Enabled = True
    cmb_organisms.Enabled = True
    txt_c_report.Enabled = True
    txt_c_comment.Enabled = True
    grid_c_anti.Enabled = True
    txt_c_1.Enabled = True
    txt_c_modflg = True
    
    'for addendum
    If USRNAME_RS("ADDENDUM") = "YES" Then
        txt_addendum.Enabled = True
        cmb_addendumby.Enabled = True
    Else
        txt_addendum.Enabled = False
        cmb_addendumby.Enabled = False
    End If
    
    Exit Sub
    
MSG:
    MsgBox Err.Description
End Sub
Private Sub disable_field()
'for panel1(bacteriology)
    txt_epith.Enabled = False
    txt_pus.Enabled = False
    txt_other.Enabled = False
    txt_koh.Enabled = False
    
    '***** Shraddha 30/06/09 *****
    'txt_organ1.Enabled = False
    'TXT_ORGAN2.Enabled = False
    'TXT_ORGAN3.Enabled = False
    
    cmb_org1.Enabled = False
    cmb_org2.Enabled = False
    cmb_org3.Enabled = False
    '*****************************
    
    txt_1.Enabled = False
    txt_2.Enabled = False
    MSFlexGrid1.Enabled = False 'Shraddha Pandit 29/10/08
    txt_test.Enabled = False
    txt_MCcomment.Enabled = False
    txt_testresult.Enabled = False
    text1.Visible = False
    txt_btest.Enabled = False
    txt_bresult.Enabled = False
    '*******
    
    txt_3.Enabled = False
    txt_comment.Enabled = False
    txt_report.Enabled = False
'for panel2(urine and stool)
    txt_urinecolour.Enabled = False
    txt_gravity.Enabled = False
    
    cmb_blood.Enabled = False
    cmb_glucose.Enabled = False
    cmb_ph.Enabled = False
    cmb_ketones.Enabled = False
    cmb_proteins.Enabled = False
    cmb_uroblngn.Enabled = False
    
    cmb_nitrite.Enabled = False 'Shraddha 21/03/11
    
    cmb_bilirubin.Enabled = False
    
    txt_ph.Enabled = False
    txt_glucose.Enabled = False
    txt_blood.Enabled = False
    txt_casts.Enabled = False
    txt_urinother.Enabled = False
    txt_urinecomment.Enabled = False
    txt_bacteria.Enabled = False
    txt_wbc.Enabled = False
    txt_ketones.Enabled = False
    txt_bilirubin.Enabled = False
    txt_rbc.Enabled = False
    'txt_MicroExam.Enabled = False
    txt_protins.Enabled = False
    txt_uro.Enabled = False
    txt_squamous.Enabled = False
    txt_stoolcolour.Enabled = False
    txt_consistency.Enabled = False
    txt_puscells.Enabled = False
    txt_stoolrbc.Enabled = False
    txt_stoolother.Enabled = False
    txt_stoolcomments.Enabled = False
    txt_stoolcolour.Enabled = False
    txt_mucus.Enabled = False
    txt_sblood.Enabled = False
    txt_cysts.Enabled = False
    txt_microphages.Enabled = False
    txt_stoolrbc.Enabled = False
    txt_Ova.Enabled = False
    txt_occult.Enabled = False
    txt_puscells.Enabled = False
    txt_fat.Enabled = False
    txt_starch.Enabled = False
    txt_larvae.Enabled = False
    txt_parasties.Enabled = False
        
        'for panel3(for other)
    'txt_refer_service.Enabled = False
    txt_test_perform.Enabled = False
    txt_result.Enabled = False
'for all combo
    cmb_gnb.Enabled = False
    cmb_gnc.Enabled = False
    cmb_gpb.Enabled = False
    cmb_gpc.Enabled = False
    cmb_india.Enabled = False
    txt_ziehl.Enabled = False
    cmb_patho.Enabled = False
 'for panel4(myco)
    cmb_stain.Enabled = False
    cmb_organisms.Enabled = False
    txt_c_report.Enabled = False
    txt_c_comment.Enabled = False
    grid_c_anti.Enabled = False
End Sub

Private Sub get_myco()
On Error GoTo MSG
    Dim ST1 As String
    ST1 = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
    
    sel_submodali_qr(0) = Trim(Mid(cmb_modality.Text, 1, InStr(cmb_modality.Text, "$") - 1))
    Set sel_submodali_rs = sel_submodali_qr.OpenResultset(1, 3)
    If Not sel_submodali_rs.EOF Then
        cmb_organisms.clear
        While Not sel_submodali_rs.EOF
            cmb_organisms.AddItem Trim(sel_submodali_rs(0))
            sel_submodali_rs.MoveNext
        Wend
        sel_submodali_rs.Close
     Else
        MsgBox "Sorry No Organism Isolated is Available,Use Master Screen For Entring  ", vbOKOnly
        sel_submodali_rs.Close
        Exit Sub
     End If
     Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub GETUSERLIST()
'On Error GoTo MSG
'    Set sel_ENTER_rs = sel_ENTER_qr.OpenResultset(1, 3)
'    If sel_ENTER_rs.RowCount > 0 Then
'        cmb_enteryby.clear
'            While Not sel_ENTER_rs.EOF
'                cmb_enteryby.AddItem Trim(sel_ENTER_rs("CCNAME"))
'                sel_ENTER_rs.MoveNext
'            Wend
'        sel_ENTER_rs.Close
'    Else
'        sel_ENTER_rs.Close
'    End If
    'INSERT
'    For i = 0 To (cmb_enteryby.ListCount - 1)
'        If Trim(cmb_enteryby.List(i)) = "Mrs. Kalpana Varadkar" Then
'                cmb_enteryby.ListIndex = i
'        End If
'    Next i
'    Exit Sub
'MSG:
'    MsgBox Err.Description
End Sub

Private Sub GETFINALLIST()
'On Error GoTo MSG
'    Set sel_FINAL_rs = sel_FINAL_qr.OpenResultset(1, 3)
'    If sel_FINAL_rs.RowCount > 0 Then
'        cmb_final.clear
'            While Not sel_FINAL_rs.EOF
'                cmb_final.AddItem Trim(sel_FINAL_rs("CCNAME"))
'                sel_FINAL_rs.MoveNext
'            Wend
'            sel_FINAL_rs.Close
'    Else
'        sel_FINAL_rs.Close
'    End If
'    'INSERT
'    For i = 0 To (cmb_final.ListCount - 1)
'        If Trim(cmb_final.List(i)) = "Dr.Rohini Kelkar" Then
'                cmb_final.ListIndex = i
'        End If
'    Next i
'    Exit Sub
'
'MSG:
'    MsgBox Err.Description
End Sub

Private Sub insert_urine()
    'insert in urine table DMURE
            ins_4_qr(0) = mod_reqno
            ins_4_qr(1) = Trim(txt_urinecolour.Text)
            ins_4_qr(2) = Trim(txt_gravity.Text)
            ins_4_qr(3) = Trim(cmb_ph.Text) 'Trim(txt_ph.Text)
            ins_4_qr(4) = Trim(cmb_proteins.Text) 'Trim(txt_protins.Text)
            ins_4_qr(5) = Trim(cmb_glucose.Text) 'Trim(txt_glucose.Text)
            ins_4_qr(6) = Trim(cmb_ketones.Text) 'Trim(txt_ketones.Text)
            ins_4_qr(7) = Trim(cmb_bilirubin.Text)  'Trim(txt_bilirubin.Text)
            ins_4_qr(8) = Trim(cmb_uroblngn.Text) 'Trim(txt_uro.Text)
            
            
            
            ins_4_qr(9) = Trim(cmb_blood.Text) 'Trim(txt_blood.Text)
            'ins_4_qr(10) = Trim(txt_MicroExam.Text)
            ins_4_qr(10) = Null
            ins_4_qr(11) = Trim(txt_wbc.Text)
            ins_4_qr(12) = Trim(txt_rbc.Text)
            ins_4_qr(13) = Trim(txt_squamous.Text)
            ins_4_qr(14) = Trim(txt_bacteria.Text)
            ins_4_qr(15) = Trim(txt_casts.Text)
            If Len(txt_urinother.Text) > 0 Then
                ins_4_qr(16) = Null
                ins_4_qr(16) = Trim(txt_urinother.Text)
            Else
                ins_4_qr(16) = Null
                ins_4_qr(16) = " "
            End If
            If Len(txt_urinecomment.Text) > 0 Then
                ins_4_qr(17) = Null
                ins_4_qr(17) = Trim(txt_urinecomment.Text)
            Else
                ins_4_qr(17) = Null
                ins_4_qr(17) = " "
            End If
            '***** Shraddha 21/03/11 *****
            ins_4_qr(18) = Trim(cmb_nitrite.Text)
            '*****************************
            ins_4_qr.Execute
'            If Not (frm_stoolexam.Enabled And frm_urineexam.Enabled) Then
'                ins_8_qr(0) = mod_reqno
'                ins_8_qr.Execute
'            End If
        update_rislib
        update_rrequisitn

Exit Sub
MSG:
    MsgBox Err.Description
End Sub
Private Sub insert_stool()
On Error GoTo MSG
            ins_5_qr(0) = mod_reqno
            ins_5_qr(1) = Trim(txt_stoolcolour.Text)
            ins_5_qr(2) = Trim(txt_consistency.Text)
            ins_5_qr(3) = Trim(txt_mucus.Text)
            ins_5_qr(4) = Trim(txt_sblood.Text)
            ins_5_qr(5) = Trim(txt_parasties.Text)
            ins_5_qr(6) = Trim(txt_occult.Text)
            ins_5_qr(7) = Trim(txt_Ova.Text)
            ins_5_qr(8) = Trim(txt_cysts.Text)
            ins_5_qr(9) = Trim(txt_larvae.Text)
            ins_5_qr(10) = Trim(txt_puscells.Text)
            ins_5_qr(11) = Trim(txt_stoolrbc.Text)
            ins_5_qr(12) = Trim(txt_microphages.Text)
            ins_5_qr(13) = Trim(txt_starch.Text)
            ins_5_qr(14) = Trim(txt_fat.Text)
            ins_5_qr(15) = Trim(txt_stoolother.Text)
            If Len(txt_stoolcomments.Text) > 0 Then
                ins_5_qr(16) = Null
                ins_5_qr(16) = txt_stoolcomments.Text
            Else
                ins_5_qr(16) = Null
                ins_5_qr(16) = " "
            End If
            ins_5_qr.Execute
            
'            If Not (frm_stoolexam.Enabled And frm_urineexam.Enabled) Then
'                ins_7_qr(0) = mod_reqno
'                ins_7_qr.Execute
'            End If
            update_rislib
    update_rrequisitn
Exit Sub
MSG:
    MsgBox Err.Description
End Sub
Private Sub insert_myco()
On Error GoTo MSG
            ins_6_qr(0) = mod_reqno
            ins_6_qr(1) = cmb_stain.Text
            ins_6_qr(2) = cmb_organisms.Text
            If Len(txt_c_report.Text) > 0 Then
                ins_6_qr(3) = Null
                ins_6_qr(3) = txt_c_report.Text
            Else
                ins_6_qr(3) = Null
                ins_6_qr(3) = " "
            End If
            If Len(txt_c_comment.Text) > 0 Then
                ins_6_qr(4) = Null
                ins_6_qr(4) = txt_c_comment.Text
            Else
                ins_6_qr(4) = Null
                ins_6_qr(4) = " "
            End If
            ins_6_qr.Execute
            I = grid_c_anti.Rows
            srno = 1
            If Not I = 1 Then 'srno = 1 Then
                If Not (Trim(grid_c_anti.TextMatrix(srno, 1)) = "") Then
                    While Not srno = I
                        ins_2_qr(0) = mod_reqno
                        ins_2_qr(1) = Trim(grid_c_anti.TextMatrix(srno, 0)) 'first get name of antibiotic and insert
                        ins_2_qr(2) = grid_c_anti.TextMatrix(srno, 1) 'get first value and insert
                        ins_2_qr(3) = Null
                        ins_2_qr(4) = Null
                        ins_2_qr.Execute
                        srno = srno + 1
                    Wend
                 Else
                        ins_2_qr(0) = mod_reqno
                        ins_2_qr(1) = Null 'first get name of antibiotic and insert
                        ins_2_qr(2) = Null 'get first value and insert
                        ins_2_qr(3) = Null
                        ins_2_qr(4) = Null
                        ins_2_qr.Execute
                 End If
             Else
                ''''INSERT SINGLE RECORD IN DATABASE FOR CRYSTAL REPORT JOIN
                ins_2_qr(0) = mod_reqno
                ins_2_qr(1) = Null
                ins_2_qr(2) = Null
                ins_2_qr(3) = Null
                ins_2_qr(4) = Null
                ins_2_qr.Execute
            End If
                update_rislib
                update_rrequisitn
Exit Sub
MSG:
    MsgBox Err.Description
End Sub
Private Sub insert_GZZSEROLOG()
On Error GoTo MSG
    'Dim RQ As rdoQuery
    'Set RQ = rdb.CreateQuery("InsertIntoGZZSEROLOG", str)
    insGZZqr(0) = mod_reqno 'REQNO
    insGZZqr(1) = Trim(txtSerTestPerform.Text) 'TESTPERF
    insGZZqr(2) = Trim(cmbSerResult.Text)   'Result
    insGZZqr(3) = Trim(cmbProcalci.Text) 'PROCALC
    insGZZqr(4) = Trim(cmbProCalCon.Text) 'PROCAL_CN
    insGZZqr(5) = Trim(cmbSerRange.Text) 'TESTRANGE
    insGZZqr(6) = Trim(txtInterpretation.Text) 'INTERPRETN
    insGZZqr(7) = Trim(txtMethod.Text) 'METHOD
    insGZZqr(8) = Trim(txtLimitation.Text) 'LIMITATION
    insGZZqr(9) = Trim(txtInBold.Text) 'FLDBLD
    insGZZqr(10) = "" 'FLD1
    insGZZqr.Execute
    update_rislib
    update_rrequisitn
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub
Private Sub update_GZZSEROLOG()
On Error GoTo MSG
    UpdGZZqr(0) = Trim(txtSerTestPerform.Text) 'TESTPERF
    UpdGZZqr(1) = Trim(cmbSerResult.Text)   'Result
    UpdGZZqr(2) = Trim(cmbProcalci.Text) 'PROCALC
    UpdGZZqr(3) = Trim(cmbProCalCon.Text) 'PROCAL_CN
    UpdGZZqr(4) = Trim(cmbSerRange.Text) 'TESTRANGE
    UpdGZZqr(5) = Trim(txtInterpretation.Text) 'INTERPRETN
    UpdGZZqr(6) = Trim(txtMethod.Text) 'METHOD
    UpdGZZqr(7) = Trim(txtLimitation.Text) 'LIMITATION
    UpdGZZqr(8) = Trim(txtInBold.Text) 'FLDBLD
    UpdGZZqr(9) = "" 'FLD1
    UpdGZZqr(10) = mod_reqno 'REQNO
    UpdGZZqr.Execute
MSG:
    MsgBox Err.Description
End Sub
Private Sub insert_other()
On Error GoTo MSG
ins_0_qr(0) = mod_reqno
            'If Len(txt_refer_service.Text) > 0 Then
                ins_0_qr(1) = Null
             '   ins_0_qr(1) = txt_refer_service.Text
            'Else
            '    ins_0_qr(1) = Null
            '    ins_0_qr(1) = " "
            'End If
            If Len(txt_test_perform.Text) > 0 Then
                ins_0_qr(2) = Null
                ins_0_qr(2) = Trim(txt_test_perform.Text)
            Else
                ins_0_qr(2) = Null
                ins_0_qr(2) = " "
            End If
            If Len(txt_result.Text) > 0 Then
                ins_0_qr(3) = Null
                ins_0_qr(3) = txt_result.Text
            Else
                ins_0_qr(3) = Null
                ins_0_qr(3) = " "
            End If
            ins_0_qr.Execute
            update_rislib
update_rrequisitn

Exit Sub
MSG:
    MsgBox Err.Description
End Sub
Private Sub UPDATE_DMBAANTI()
                     
''''' Done By Kanchan
 ''''' Check for data in the database
 
                    sel_forupd2_qr(0) = mod_reqno
                    Set sel_forupd2_rs = sel_forupd2_qr.OpenResultset(1, 3)
                                     
                    If Not sel_forupd2_rs.EOF Then
                    
                        row1 = 0
                        row1 = grid_dmbanti.Rows - 1
    
                            If Len(Trim(grid_dmbanti.TextMatrix(row1, 0))) > 0 Then
                                row1 = grid_dmbanti.Rows - 1
                            Else
                                row1 = grid_dmbanti.Rows - 2
                            End If
                        
                        srno = 1
                        For I = 1 To row1
                                '******* Shraddha 30/06/09 *******
                                'If Len(Trim(txt_organ1.Text)) >= 1 Then
                                If Len(Trim(cmb_org1.Text)) >= 1 Then
                                
                                    'If Len(Trim(TXT_ORGAN2.Text)) >= 1 Then
                                    If Len(Trim(cmb_org2.Text)) >= 1 Then
                                    
                                        'If Len(Trim(TXT_ORGAN3.Text)) >= 1 Then
                                        If Len(Trim(cmb_org3.Text)) >= 1 Then
                                '*****************************************
                                            ''''''' three org
                                            If (Len(Trim(grid_dmbanti.TextMatrix(srno, 0))) > 0) Then '''''CHECK FOR ANTIBIOTIC
                                                If ((Not Trim(grid_dmbanti.TextMatrix(srno, 1)) = "") Or (Not Trim(grid_dmbanti.TextMatrix(srno, 2)) = "") Or (Not Trim(grid_dmbanti.TextMatrix(srno, 3)) = "")) Then
                                                    
                                                        sel_formod2_qr(0) = mod_reqno
                                                        sel_formod2_qr(1) = Trim(grid_dmbanti.TextMatrix(srno, 0))
                                                        If sel_formod2_flg = False Then
                                                            Set sel_formod2_rs = sel_formod2_qr.OpenResultset(1, 3)
                                                            sel_formod2_flg = True
                                                        Else
                                                            sel_formod2_rs.MoreResults
                                                            sel_formod2_rs.Requery
                                                        End If
                                    
                                                        If sel_formod2_rs.RowCount > 0 Then
                                                            upd_2_qr(0) = Trim(grid_dmbanti.TextMatrix(srno, 1)) 'get first value and insert
                                                            upd_2_qr(1) = Trim(grid_dmbanti.TextMatrix(srno, 2)) 'get second value and insert"
                                                            upd_2_qr(2) = Trim(grid_dmbanti.TextMatrix(srno, 3)) 'get third value and insert
                                                            upd_2_qr(3) = mod_reqno
                                                            upd_2_qr(4) = Trim(grid_dmbanti.TextMatrix(srno, 0)) 'first get name of antibiotic and insert
                                                            upd_2_qr.Execute
                                                        Else
                                    
                                                        'reqno,antibiotic,antibio1,antibio2,antibio3
                                            
                                                            ins_2_qr(0) = Trim(mod_reqno)
                                                            ins_2_qr(1) = Trim(grid_dmbanti.TextMatrix(srno, 0))
                                                            ins_2_qr(2) = Trim(grid_dmbanti.TextMatrix(srno, 1))
                                                            ins_2_qr(3) = Trim(grid_dmbanti.TextMatrix(srno, 2))
                                                            ins_2_qr(4) = Trim(grid_dmbanti.TextMatrix(srno, 3))
                                                            ins_2_qr.Execute
                                                        End If
                                                   
                                                Else
                                                    ''''delete from table
                                                    
                                                      del_dmbaanti_qr1(0) = mod_reqno
                                                      del_dmbaanti_qr1(1) = Trim(grid_dmbanti.TextMatrix(srno, 0))
                                                      del_dmbaanti_qr1.Execute
                                      
                                                End If
                                             End If ''''
                                        Else
                                            ''''''' two org
                                                If (Len(Trim(grid_dmbanti.TextMatrix(srno, 0))) > 0) Then '''''CHECK FOR ANTIBIOTIC
                                                    If ((Not Trim(grid_dmbanti.TextMatrix(srno, 1)) = "") Or (Not Trim(grid_dmbanti.TextMatrix(srno, 2)) = "")) Then
                                                    
                                                        sel_formod2_qr(0) = mod_reqno
                                                        sel_formod2_qr(1) = Trim(grid_dmbanti.TextMatrix(srno, 0))
                                                        If sel_formod2_flg = False Then
                                                            Set sel_formod2_rs = sel_formod2_qr.OpenResultset(1, 3)
                                                            sel_formod2_flg = True
                                                        Else
                                                            sel_formod2_rs.MoreResults
                                                            sel_formod2_rs.Requery
                                                        End If
                                    
                                                        If sel_formod2_rs.RowCount > 0 Then
                                                            upd_2_qr(0) = Trim(grid_dmbanti.TextMatrix(srno, 1)) 'get first value and insert
                                                            upd_2_qr(1) = Trim(grid_dmbanti.TextMatrix(srno, 2)) 'get second value and insert"
                                                            upd_2_qr(2) = Null 'get third value and insert
                                                            upd_2_qr(3) = mod_reqno
                                                            upd_2_qr(4) = Trim(grid_dmbanti.TextMatrix(srno, 0)) 'first get name of antibiotic and insert
                                                            upd_2_qr.Execute
                                                        Else
                                    
                                                        'reqno,antibiotic,antibio1,antibio2,antibio3
                                            
                                                            ins_2_qr(0) = Trim(mod_reqno)
                                                            ins_2_qr(1) = Trim(grid_dmbanti.TextMatrix(srno, 0))
                                                            ins_2_qr(2) = Trim(grid_dmbanti.TextMatrix(srno, 1))
                                                            ins_2_qr(3) = Trim(grid_dmbanti.TextMatrix(srno, 2))
                                                            ins_2_qr(4) = Null
                                                            ins_2_qr.Execute
                                                        End If
                                        
                                                Else
                                                    ''''delete from table
                                                      del_dmbaanti_qr1(0) = mod_reqno
                                                      del_dmbaanti_qr1(1) = Trim(grid_dmbanti.TextMatrix(srno, 0))
                                                      del_dmbaanti_qr1.Execute
                                      
                                                End If
                                            
                                            End If '''''
                                        End If
                                    Else
                                            '''''''one org
                                        If (Len(Trim(grid_dmbanti.TextMatrix(srno, 0))) > 0) Then '''''CHECK FOR ANTIBIOTIC
                                            If ((Not Trim(grid_dmbanti.TextMatrix(srno, 1)) = "")) Then
                                                    
                                                        sel_formod2_qr(0) = mod_reqno
                                                        sel_formod2_qr(1) = Trim(grid_dmbanti.TextMatrix(srno, 0))
                                                        If sel_formod2_flg = False Then
                                                            Set sel_formod2_rs = sel_formod2_qr.OpenResultset(1, 3)
                                                            sel_formod2_flg = True
                                                        Else
                                                            sel_formod2_rs.MoreResults
                                                            sel_formod2_rs.Requery
                                                        End If
                                    
                                                        If sel_formod2_rs.RowCount > 0 Then
                                                            upd_2_qr(0) = Trim(grid_dmbanti.TextMatrix(srno, 1)) 'get first value and insert
                                                            upd_2_qr(1) = Null 'get second value and insert"
                                                            upd_2_qr(2) = Null 'get third value and insert
                                                            upd_2_qr(3) = mod_reqno
                                                            upd_2_qr(4) = Trim(grid_dmbanti.TextMatrix(srno, 0)) 'first get name of antibiotic and insert
                                                            upd_2_qr.Execute
                                                        Else
                                    
                                                        'reqno,antibiotic,antibio1,antibio2,antibio3
                                            
                                                            ins_2_qr(0) = Trim(mod_reqno)
                                                            ins_2_qr(1) = Trim(grid_dmbanti.TextMatrix(srno, 0))
                                                            ins_2_qr(2) = Trim(grid_dmbanti.TextMatrix(srno, 1))
                                                            ins_2_qr(3) = Null
                                                            ins_2_qr(4) = Null
                                                            ins_2_qr.Execute
                                                        End If
                                        
                                                Else
                                                    ''''delete from table
                                                      del_dmbaanti_qr1(0) = mod_reqno
                                                      del_dmbaanti_qr1(1) = Trim(grid_dmbanti.TextMatrix(srno, 0))
                                                      del_dmbaanti_qr1.Execute
                                      
                                                End If
                                           End If '''''
                                    End If
                                End If
                             
                            srno = srno + 1
                            
                        Next I
                        
                    Else
                        INSERT_DMBANTI
                    End If

End Sub

Private Sub INSERT_DMBANTI()

'''' Done by KANCHAN
'******************* Shraddha 30/06/09 ***********************
'If Len(txt_organ1.Text) >= 1 Or Len(TXT_ORGAN2.Text) >= 1 Or Len(TXT_ORGAN3.Text) >= 1 Then
If Len(cmb_org1.Text) >= 1 Or Len(cmb_org2.Text) >= 1 Or Len(cmb_org3.Text) >= 1 Then
'*************************************************************************
    I = grid_dmbanti.Rows - 1
    
    If Len(Trim(grid_dmbanti.TextMatrix(I, 0))) > 0 Then
        I = grid_dmbanti.Rows
    Else
        I = grid_dmbanti.Rows - 1
    End If
    
'    i = grid_dmbanti.Rows
    srno = 1
    
    If I = 1 Then
        
            ins_2_qr(0) = Trim(mod_reqno)
            ins_2_qr(1) = Null
            ins_2_qr(2) = Null
            ins_2_qr(3) = Null
            ins_2_qr(4) = Null
            ins_2_qr.Execute
    
    End If
    While Not srno = I
    
        '******** Shraddha 30/06/09 *******
        'If Len(txt_organ1.Text) >= 1 Then
        If Len(cmb_org1.Text) >= 1 Then
        
            'If Len(TXT_ORGAN2.Text) >= 1 Then
            If Len(cmb_org2.Text) >= 1 Then
            
                'If Len(TXT_ORGAN3.Text) >= 1 Then
                If Len(cmb_org3.Text) >= 1 Then
        '**********************************
                    '''THREE ORG
                    If (Len(Trim(grid_dmbanti.TextMatrix(srno, 0))) > 0) Then
                        If (Len(Trim(grid_dmbanti.TextMatrix(srno, 1))) > 0 Or Len(Trim(grid_dmbanti.TextMatrix(srno, 2))) > 0 Or Len(Trim(grid_dmbanti.TextMatrix(srno, 3))) > 0) Then
                    
                            ins_2_qr(0) = Trim(mod_reqno)
                            ins_2_qr(1) = Trim(grid_dmbanti.TextMatrix(srno, 0))
                            ins_2_qr(2) = Trim(grid_dmbanti.TextMatrix(srno, 1))
                            ins_2_qr(3) = Trim(grid_dmbanti.TextMatrix(srno, 2))
                            ins_2_qr(4) = Trim(grid_dmbanti.TextMatrix(srno, 3))
                            ins_2_qr.Execute
                        End If
                    
                    End If
                Else
                    ''''TWO ORG
                    If (Len(Trim(grid_dmbanti.TextMatrix(srno, 0))) > 0) Then
                        If (Len(Trim(grid_dmbanti.TextMatrix(srno, 1))) > 0 Or Len(Trim(grid_dmbanti.TextMatrix(srno, 2))) > 0) Then
                            ins_2_qr(0) = Trim(mod_reqno)
                            ins_2_qr(1) = Trim(grid_dmbanti.TextMatrix(srno, 0))
                            ins_2_qr(2) = Trim(grid_dmbanti.TextMatrix(srno, 1))
                            ins_2_qr(3) = Trim(grid_dmbanti.TextMatrix(srno, 2))
                            ins_2_qr(4) = Null
                            ins_2_qr.Execute
                    
                        End If
                    End If
                 End If
            Else
                
                   ''ONE ORG
                    If (Len(Trim(grid_dmbanti.TextMatrix(srno, 0))) > 0) Then
                        If (Len(Trim(grid_dmbanti.TextMatrix(srno, 1))) > 0) Then
                            ins_2_qr(0) = Trim(mod_reqno)
                            ins_2_qr(1) = Trim(grid_dmbanti.TextMatrix(srno, 0))
                            ins_2_qr(2) = Trim(grid_dmbanti.TextMatrix(srno, 1))
                            ins_2_qr(3) = Null
                            ins_2_qr(4) = Null
                            ins_2_qr.Execute
                        End If
                    End If
            End If
        End If
    
        srno = srno + 1
    Wend

Else
        
        'MsgBox ("Atlest 1 Organism is needed....!"), vbOKOnly
        'txt_1.SetFocus
               
'      ins_2_qr(0) = Trim(mod_reqno)
'      ins_2_qr(1) = Null
'      ins_2_qr(2) = Null
'      ins_2_qr(3) = Null
'      ins_2_qr(4) = Null
'      ins_2_qr.Execute
End If

'ins_2_qr.Execute
End Sub

Private Sub insert_bacteriology()

'select dmoprismer
    Dim rs_dmoprismer As rdoResultset
    Dim del_dmoprismer As String
    Set rs_dmoprismer = rdb.OpenResultset("select * from dislib.dmoprismer where reqno ='" & Trim(mod_reqno) & "'", 1, 1)
    If rs_dmoprismer.EOF = False Then
    del_dmoprismer = "delete from dislib.dmoprismer where reqno ='" & Trim(mod_reqno) & "'"
    rdb.Execute del_dmoprismer
    End If
    '''upto here



'delete dmoprismer

'On Error GoTo MSG
            ins_1_qr(0) = mod_reqno
            ins_1_qr(1) = Trim(txt_epith.Text)
            ins_1_qr(2) = Trim(txt_pus.Text)
            ins_1_qr(3) = Trim(txt_other.Text)
            'ins_1_qr(4) = txt_ziehl.Text
            ins_1_qr(4) = Trim(txt_ziehl.Text)
            ins_1_qr(5) = Trim(txt_koh.Text)
            ins_1_qr(6) = Trim(cmb_gpc.Text)
            ins_1_qr(7) = Trim(cmb_gnc.Text)
            ins_1_qr(8) = Trim(cmb_gpb.Text)
            ins_1_qr(9) = Trim(cmb_gnb.Text)
            ins_1_qr(10) = Trim(cmb_india.Text)
            
            '************ Shraddha 30/06/09 ***********
'            If Len(txt_organ1.Text) > 0 Then
'                ins_1_qr(11) = Null
'                ins_1_qr(11) = Trim(txt_organ1.Text)
'            Else
'                ins_1_qr(11) = Null
'                ins_1_qr(11) = " "
'            End If
            
            If Len(cmb_org1.Text) > 0 Then
                ins_1_qr(11) = Null
                ins_1_qr(11) = Trim(cmb_org1.Text)
            Else
                ins_1_qr(11) = Null
                ins_1_qr(11) = " "
            End If
            
            
'            If Len(TXT_ORGAN2.Text) > 0 Then
'                ins_1_qr(12) = Null
'                ins_1_qr(12) = Trim(TXT_ORGAN2.Text)
'            Else
'                ins_1_qr(12) = Null
'                ins_1_qr(12) = " "
'            End If
            
            If Len(cmb_org2.Text) > 0 Then
                ins_1_qr(12) = Null
                ins_1_qr(12) = Trim(cmb_org2.Text)
            Else
                ins_1_qr(12) = Null
                ins_1_qr(12) = " "
            End If
            
'            If Len(TXT_ORGAN3.Text) > 0 Then
'                ins_1_qr(13) = Null
'                ins_1_qr(13) = Trim(TXT_ORGAN3.Text)
'            Else
'                ins_1_qr(13) = Null
'                ins_1_qr(13) = " "
'            End If

            If Len(cmb_org3.Text) > 0 Then
                ins_1_qr(13) = Null
                ins_1_qr(13) = Trim(cmb_org3.Text)
            Else
                ins_1_qr(13) = Null
                ins_1_qr(13) = " "
            End If
            
            '*******************************************
            
            If cmb_patho.Text = "Yes" Then
                ins_1_qr(14) = Null
            Else
                If Len(txt_report.Text) > 0 Then
                    ins_1_qr(14) = Null
                    ins_1_qr(14) = Trim(txt_report.Text)
                Else
                    ins_1_qr(14) = Null
                    ins_1_qr(14) = " "
                End If
            End If
            If Len(txt_comment.Text) > 0 Then
                'ins_1_qr(15) = Null 'Shraddha Pandit 17/11/2008
                ins_1_qr(15) = Trim(txt_comment.Text)
            Else
                ins_1_qr(15) = Null
                ins_1_qr(15) = " "
            End If
            ins_1_qr.Execute
            'MsgBox "Record successfully saved" 'Insert Commented on 05/06/09 by Shraddha
            
   INSERT_DMBANTI
            
'''''''''COMMENTED BY KANCHAN
            
'            'for inserting effect of antibiotics for modality 1
'            'here dept,srno is used to get total no. of rows
'            'CHECK HOW MANY ORGANISMS
'
'
'
'            If Len(txt_organ1.Text) > 1 Or Len(TXT_ORGAN2.Text) > 1 Or Len(TXT_ORGAN3.Text) > 1 Then
'            'THIS UPPER IF IS FOR CRYSTAL REPORT JOIN IF NON ANTIBIOTIC IS SELECTED THEN INSERT ONLY REQNO IN TABLE.
'                If Len(txt_organ1.Text) > 1 Then
'                    If Len(TXT_ORGAN2.Text) > 1 Then
'                        If Len(TXT_ORGAN3.Text) > 1 Then
'                        'three organisms
'                            i = grid_anti.Rows
'                            srno = 1
'                            While Not srno = i
'                            'check for all the three entry field they should not be empty
'                               ' If Not ((Trim(grid_anti.TextMatrix(SRNO, 1)) Is Null) Or (Trim(grid_anti.TextMatrix(SRNO, 2)) Is Null) Or (Trim(grid_anti.TextMatrix(SRNO, 3)) Is Null)) Then
'                                If ((Not Trim(grid_anti.TextMatrix(srno, 1)) = "") Or (Not Trim(grid_anti.TextMatrix(srno, 2)) = "") Or (Not Trim(grid_anti.TextMatrix(srno, 3)) = "")) Then
'                                    ins_2_qr(0) = mod_reqno
'                                    ins_2_qr(1) = Trim(grid_anti.TextMatrix(srno, 0)) 'first get name of antibiotic and insert
'                                    ins_2_qr(2) = grid_anti.TextMatrix(srno, 1) 'get first value and insert
'                                    ins_2_qr(3) = grid_anti.TextMatrix(srno, 2) 'get second value and insert"
'                                    ins_2_qr(4) = grid_anti.TextMatrix(srno, 3) 'get third value and insert
'                                    ins_2_qr.Execute
'                                End If
'                                srno = srno + 1
'                            Wend
'                        Else
'                        'two organisms
'                            i = grid_anti.Rows
'                            srno = 1
'                            While Not srno = i
'                            If ((Not Trim(grid_anti.TextMatrix(srno, 1)) = "") Or (Not Trim(grid_anti.TextMatrix(srno, 2)) = "")) Then
'                                ins_2_qr(0) = mod_reqno
'                                ins_2_qr(1) = Trim(grid_anti.TextMatrix(srno, 0)) 'first get name of antibiotic and insert
'                                ins_2_qr(2) = grid_anti.TextMatrix(srno, 1) 'get first value and insert
'                                ins_2_qr(3) = grid_anti.TextMatrix(srno, 2) 'get second value and insert"
'                                ins_2_qr(4) = Null
'                                ins_2_qr.Execute
'                            End If
'                                srno = srno + 1
'                            Wend
'                        End If
'                    Else
'                    'only one organisms
'                    i = grid_anti.Rows
'                        srno = 1
'                        While Not srno = i
'                        If Not (Trim(grid_anti.TextMatrix(srno, 1)) = "") Then
'                            ins_2_qr(0) = Trim(mod_reqno)
'                            ins_2_qr(1) = Trim(grid_anti.TextMatrix(srno, 0)) 'first get name of antibiotic and insert
'                            ins_2_qr(2) = grid_anti.TextMatrix(srno, 1) 'get first value and insert
'                            ins_2_qr(3) = Null
'                            ins_2_qr(4) = Null
'                            ins_2_qr.Execute
'                        End If
'                            srno = srno + 1
'                        Wend
'                   End If
'                End If
'            Else
'                ins_2_qr(0) = Trim(mod_reqno)
'                ins_2_qr(1) = Null
'                ins_2_qr(2) = Null
'                ins_2_qr(3) = Null
'                ins_2_qr(4) = Null
'                ins_2_qr.Execute
'            End If

            update_rislib
    update_rrequisitn
Exit Sub
MSG:
    MsgBox Err.Description
End Sub
Private Sub update_rrequisitn()
On Error GoTo MSG
    upd_11_qr(0) = MaskEdBox1.FormattedText
    upd_11_qr(1) = mod_reqno
    upd_11_qr.Execute
   ' MsgBox "Record Has Been Saved Successfully", vbOKOnly
    Label27.Caption = "Record Has Been Saved Successfully"
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub update_rislib()
On Error GoTo MSG
  'MODIFIED BY MRUNAL
    Dim GS_QR As rdoQuery
        Dim GS_RS As rdoResultset
        str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO = ? "
        Set GS_QR = rdb.CreateQuery("", str)
        GS_QR(0) = RQN
        Set GS_RS = GS_QR.OpenResultset(1, 3)
        If GS_RS.RowCount > 0 Then
        If InStr(GS_RS("REQNSTAT"), "E") Then
                    UPDRTIMEC_QR1(0) = Format(txt_date.Text, "YYYY-MM-DD")
    UPDRTIMEC_QR1(1) = Format(txt_time.Text, "HH:MM:SS")
                    
    UPDRTIMEC_QR1(2) = mod_reqno
        UPDRTIMEC_QR1.Execute
            End If
                   If InStr(GS_RS("REQNSTAT"), "T") Then
           UPDRTIMEC_QR2(0) = Format(txt_date.Text, "YYYY-MM-DD")
    UPDRTIMEC_QR2(1) = Format(txt_time.Text, "HH:MM:SS")
                    
    UPDRTIMEC_QR2(2) = mod_reqno
        UPDRTIMEC_QR2.Execute
            End If
            
            End If
            
    
        ' END ---------
    upd_8_qr(0) = cmb_enteryby.Text
    upd_8_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    upd_8_qr(2) = cmb_enteryby.Text
    upd_8_qr(3) = "T"
'    upd_8_qr(4) = cmb_final.Text
'    upd_8_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
upd_8_qr(4) = Null
upd_8_qr(5) = Null
    upd_8_qr(6) = Trim(mod_reqno)
    
    cmd_save.Enabled = False 'MODIFIED BY MRUNAL 23_2
    cmd_commit.Enabled = True 'MODIFIED BY MRUNAL 23_2
    upd_8_qr.Execute
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub GET_ENTER_BY()
Dim ENTR_BY As String
On Error GoTo MSG
    sel_ris_qr(0) = mod_reqno
    Set sel_ris_rs = sel_ris_qr.OpenResultset(1, 3)
    If Not sel_ris_rs.EOF Then
        While Not sel_ris_rs.EOF
            'cmb_enteryby.AddItem sel_ris_rs("ENTERBY") & ""
            cmb_enteryby.Text = sel_ris_rs("ENTERBY") & ""
            If cmb_final.Text = "" Then
            cmb_final.Text = sel_ris_rs("freportby") & ""
            End If
            'ENTR_BY = ""
            'ENTR_BY = Mid(sel_ris_rs("freportby"), 4)
            'cmb_final.Text = Trim(ENTR_BY)       ' sel_ris_rs("freportby") & ""
            cmb_enteryby.Enabled = False
            cmb_final.Enabled = False
        sel_ris_rs.MoveNext
        Wend
        sel_ris_rs.Close
    Else
        sel_ris_rs.Close
    End If
    Exit Sub
MSG:
    MsgBox Err.Description

End Sub
Private Sub UPDATE_URINE()
On Error GoTo MSG
    upd_4_qr(0) = txt_urinecolour.Text
    upd_4_qr(1) = txt_gravity.Text
    upd_4_qr(2) = cmb_ph.Text 'txt_ph.Text
    upd_4_qr(3) = cmb_proteins.Text 'txt_protins.Text
    upd_4_qr(4) = cmb_glucose.Text 'txt_glucose.Text
    upd_4_qr(5) = cmb_ketones.Text 'txt_ketones.Text
    upd_4_qr(6) = cmb_bilirubin.Text 'txt_bilirubin.Text
    upd_4_qr(7) = cmb_uroblngn.Text 'txt_uro.Text
    upd_4_qr(8) = cmb_blood.Text 'txt_blood.Text
    'upd_4_qr(9) = txt_MicroExam.Text
    upd_4_qr(9) = Null
    upd_4_qr(10) = txt_wbc.Text
    upd_4_qr(11) = txt_rbc.Text
    upd_4_qr(12) = txt_squamous.Text
    upd_4_qr(13) = txt_bacteria.Text
    upd_4_qr(14) = txt_casts.Text
    If Len(txt_urinother.Text) > 0 Then
        upd_4_qr(15) = Null
        upd_4_qr(15) = txt_urinother.Text
    Else
        upd_4_qr(15) = Null
        upd_4_qr(15) = " "
    End If
    If Len(txt_urinecomment.Text) > 0 Then
        upd_4_qr(16) = Null
        upd_4_qr(16) = txt_urinecomment.Text
    Else
        upd_4_qr(16) = Null
        upd_4_qr(16) = " "
    End If
    upd_4_qr(17) = Trim(cmb_nitrite.Text)
    upd_4_qr(18) = mod_reqno
    upd_4_qr.Execute
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub UPDATE_STOOL()
On Error GoTo MSG
    upd_5_qr(0) = txt_stoolcolour.Text
    upd_5_qr(1) = txt_consistency.Text
    upd_5_qr(2) = txt_mucus.Text
    upd_5_qr(3) = txt_sblood.Text
    upd_5_qr(4) = txt_parasties.Text
    upd_5_qr(5) = txt_occult.Text
    upd_5_qr(6) = txt_Ova.Text
    upd_5_qr(7) = txt_cysts.Text
    upd_5_qr(8) = txt_larvae.Text
    upd_5_qr(9) = txt_puscells.Text
    upd_5_qr(10) = txt_stoolrbc.Text
    upd_5_qr(11) = txt_microphages.Text
    upd_5_qr(12) = txt_starch.Text
    upd_5_qr(13) = txt_fat.Text
    upd_5_qr(14) = txt_stoolother.Text
    If Len(txt_stoolcomments.Text) > 0 Then
        upd_5_qr(15) = Null
        upd_5_qr(15) = txt_stoolcomments.Text
    Else
        upd_5_qr(15) = Null
        upd_5_qr(15) = " "
    End If
    upd_5_qr(16) = mod_reqno
    upd_5_qr.Execute
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub
Private Sub STAFF()
On Error GoTo MSG
If new_flag Then
            ''******ICU****************
            'VAR_REQNNO = req_chglen(cmb_reqno.Text)
            If SELICU_FLAG = False Then
               SELICU_QR(0) = mod_reqno ' VAR_REQNNO
               Set SELICU_RS = SELICU_QR.OpenResultset(1, 3)
               SELICU_FLAG = True
            Else
               SELICU_QR(0) = mod_reqno 'VAR_REQNNO
               SELICU_RS.MoreResults
               SELICU_RS.Requery
            End If
            If Not SELICU_RS.EOF Then
               If Trim(SELICU_RS("ICDS2")) = "YES" Then
                  lblicu.Visible = True
               Else
                  lblicu.Visible = False
               End If
            End If
            ''******ICU****************
    
        SEL_dreq_staff_qr(0) = mod_reqno
        SEL_dreq_staff_qr(1) = MaskEdBox1.FormattedText
        Set SEL_dreq_staff_rs = SEL_dreq_staff_qr.OpenResultset(1, 3)
        If SEL_dreq_staff_rs.RowCount > 0 Then
            txt_name = SEL_dreq_staff_rs("PNAME")
            txt_sex = SEL_dreq_staff_rs("SEX")
            txt_age = SEL_dreq_staff_rs("AGE")
            If SEL_dreq_staff_rs("RELATION") <> "STAFF" Then
                txt_status = "DP"
            Else
                txt_status = "ST"
            End If
        End If
    REFDR_QR(0) = MaskEdBox1.FormattedText
    REFDR_QR(1) = REQNO
    Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
    If REFDR_RS.RowCount > 0 Then
        txt_refdoc.Text = REFDR_RS("REFBYDOC")
    End If
    sel_bed_qr(0) = MaskEdBox1.FormattedText
    Set sel_bed_rs = sel_bed_qr.OpenResultset(1, 3)
        If sel_bed_rs.RowCount > 0 Then
            txt_ward.Visible = True
            txt_bedno.Visible = True
            lbl_ward.Visible = True
           
            txt_ward.Text = sel_bed_rs("WARD_NAME")
            txt_bedno.Text = sel_bed_rs("bed_no")
            sel_bed_rs.Close
        Else
            txt_status.Text = "OP"
            txt_ward.Visible = False
            txt_bedno.Visible = False
            lbl_ward.Visible = False
          
            sel_bed_rs.Close
        End If
        INSERT
 ElseIf VIEW_FLAG Then
            ''******ICU****************
            'VAR_REQNNO = req_chglen(cmb_reqno.Text)
            If SELICU_FLAG = False Then
               SELICU_QR(0) = mod_reqno ' VAR_REQNNO
               Set SELICU_RS = SELICU_QR.OpenResultset(1, 3)
               SELICU_FLAG = True
            Else
               SELICU_QR(0) = mod_reqno 'VAR_REQNNO
               SELICU_RS.MoreResults
               SELICU_RS.Requery
            End If
            If Not SELICU_RS.EOF Then
               If Trim(SELICU_RS("ICDS2")) = "YES" Then
                  lblicu.Visible = True
               Else
                  lblicu.Visible = False
               End If
            End If
            ''******ICU****************
       
        SEL_dreq_staff_qr(0) = mod_reqno
        SEL_dreq_staff_qr(1) = MaskEdBox1.Text
        Set SEL_dreq_staff_rs = SEL_dreq_staff_qr.OpenResultset(1, 3)
        If SEL_dreq_staff_rs.RowCount > 0 Then
            txt_name = SEL_dreq_staff_rs("PNAME")
            txt_sex = SEL_dreq_staff_rs("SEX")
            txt_age = SEL_dreq_staff_rs("AGE")
            If SEL_dreq_staff_rs("RELATION") <> "STAFF" Then
                txt_status = "DP"
            Else
                txt_status = "ST"
            End If
        End If
    REFDR_QR(0) = MaskEdBox1.FormattedText
    REFDR_QR(1) = mod_reqno
    Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
    If REFDR_RS.RowCount > 0 Then
        txt_refdoc.Text = REFDR_RS("REFBYDOC")
    End If
    sel_bed_qr(0) = MaskEdBox1.FormattedText
    Set sel_bed_rs = sel_bed_qr.OpenResultset(1, 3)
        If sel_bed_rs.RowCount > 0 Then
            txt_ward.Visible = True
            txt_bedno.Visible = True
            lbl_ward.Visible = True
           
            txt_ward.Text = sel_bed_rs("WARD_NAME")
            txt_bedno.Text = sel_bed_rs("bed_no")
            sel_bed_rs.Close
        Else
            txt_status.Text = "OP"
            txt_ward.Visible = False
            txt_bedno.Visible = False
            lbl_ward.Visible = False
            
            sel_bed_rs.Close
        End If
        RETRIVE
 End If
Exit Sub
MSG:
    MsgBox Err.Description
End Sub
Private Sub INSERT()
'On Error GoTo MSG
'Shraddha 23/01/09
SEL_DMUL_QR(0) = mod_reqno
Set SEL_DMUL_RS = SEL_DMUL_QR.OpenResultset(1, 3)
enable_field
        getupperpanel
        If Trim(txt_modcode.Text) = "BA" Then 'for case no. 1(bacteriology)
            set_visible "frm_bacteria"
        ElseIf Trim(txt_modcode.Text) = "CM" Then ' for case no. 5(clinical micro)
            If Trim(CMB_PROC.Text) = "Pregnancy Test (Urinary Beta_HCG)" Then
                set_visible "frame3"
                Exit Sub
            End If
            If Trim(CMB_PROC.ListCount) > 1 Then
                set_visible "frm_stoolexam"
            ElseIf UCase(Trim(CMB_PROC.Text)) = "URINE EXAMINATION" Then
                    set_visible "frm_urineexam"
                    frm_urineexam.Enabled = True
                    frm_stoolexam.Enabled = True
               'ElseIf Trim(CMB_PROC.Text) = "STOOL EXAMINATION" Then
               
            ElseIf Not SEL_DMUL_RS.EOF And (SEL_DMUL_RS("PRO_CODE") = "G252") Then
            'Commented next and added above by Shraddha on 27/01/09
                'If SEL_DMUL_RS("PRO_CODE") = "G252" Then
                    'Next statement added by Shraddha Pandit on 23/01/09
                    set_visible "frm_others"
                    
                'End If
            Else
                    set_visible "frm_stoolexam"
                    frm_urineexam.Enabled = True
                    frm_stoolexam.Enabled = True
            End If
        ElseIf Trim(txt_modcode.Text) = "MB" Then
                
                If UCase(Trim(CMB_PROC.Text)) = "ADA LEVEL" Then
                    set_visible "frm_others"
                    Exit Sub
                Else
                    set_visible "frm_myco"
                    get_myco
                    cmb_stain.clear
                    cmb_stain.AddItem "AFB Not Seen"
                    cmb_stain.AddItem "AFB Seen"
                    cmb_stain.AddItem "AFB Seen 1+"
                    cmb_stain.AddItem "AFB Seen 2+"
                    cmb_stain.AddItem "AFB Seen 3+"
                    cmb_stain.AddItem "AFB Seen 4+"
                    cmb_stain.AddItem "Material Saliva,Please Repeat"
                End If
'        ElseIf Trim(txt_modcode.Text) = "SE" Then
'            set_visible "frmSerology"
        Else
        ' for other case
                set_visible "frm_others"
                
                txt_addendum.Enabled = False
                cmb_addendumby.Enabled = False
        End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub RETRIVE()
'On Error GoTo MSG
Dim ADDENDUM As String
'Shraddha 23/01/09
SEL_DMUL_QR(0) = mod_reqno
Set SEL_DMUL_RS = SEL_DMUL_QR.OpenResultset(1, 3)


getupperpanel
        GET_ENTER_BY
'        REFDR_QR(0) = MaskEdBox1.FormattedText
'        REFDR_QR(1) = mod_reqno
'        Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
'        If REFDR_RS.RowCount > 0 Then
'            If REFDR_RS("COMMITFLAG") = "Y" Then
'                MsgBox "Final Report is Generated, You cann't Modify the Final Report"
'                cmd_modify.Enabled = False
'                ADDENDUMFALG = True
'            End If
'            REFDR_RS.Close
'         Else
'         REFDR_RS.Close
'         End If
        'If S2 = "BA" Then '-- Shraddha Pandit 14/11/08
        If (S2 = "BA" And Frame3.Visible = False And q = False) Then
            set_visible "frm_bacteria"
            getresult_bacteriology
        ElseIf S2 = "CM" Then
            If UCase(Trim(CMB_PROC.Text)) = "PREGNANCY TEST (URINARY ?-HCG)" Then
                set_visible "frm_others"
                getresult_other
            End If
            If Trim(CMB_PROC.ListCount) > 1 Then
                set_visible "frm_stoolexam"
                getresult_urine
                getresult_stool
            ElseIf UCase(Trim(CMB_PROC.Text)) = "URINE EXAMINATION" Then 'Ucase added on 02/11/09
                set_visible "frm_urineexam"
                frm_stoolexam.Enabled = False
                frm_urineexam.Enabled = True
                getresult_urine
            ElseIf UCase(Trim(CMB_PROC.Text)) = "STOOL EXAMINATION" Or UCase(Trim(CMB_PROC.Text)) = "STOOL FOR OCCULT BLOOD" Then
                set_visible "frm_stoolexam"
                frm_urineexam.Enabled = False
                frm_stoolexam.Enabled = True
                getresult_stool
            ElseIf Not SEL_DMUL_RS.EOF Then
                If SEL_DMUL_RS("PRO_CODE") = "G252" Then
                    'Next statement added by Shraddha Pandit on 23/01/09
                    set_visible "frm_others"
                    getresult_other
                
                End If
            End If
        ElseIf S2 = "MB" Then
            If UCase(Trim(CMB_PROC.Text)) = "ADA LEVEL" Then
                set_visible "frm_others"
                getresult_other
            Else
                set_visible "frm_myco"
                get_myco
                cmb_stain.clear
                cmb_stain.AddItem "AFB Not Seen"
                cmb_stain.AddItem "AFB Seen"
                cmb_stain.AddItem "AFB Seen 1+"
                cmb_stain.AddItem "AFB Seen 2+"
                cmb_stain.AddItem "AFB Seen 3+"
                cmb_stain.AddItem "AFB Seen 4+"
                cmb_stain.AddItem "Material Saliva,Please Repeat"
                getresult_myco
            End If
        ElseIf S2 = "SE" Then
            set_visible "frmSerology"
            getresult_other
            
        ElseIf q = True Then
            set_visible "frame3"
            Dim gk
        Dim GB_QR1 As rdoQuery
        Dim GB_RS1 As rdoResultset
        str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO = ? "
        Set GB_QR1 = rdb.CreateQuery("", str)
        GB_QR1(0) = RQN
        Set GB_RS1 = GB_QR1.OpenResultset(1, 3)
        If GB_RS1.RowCount > 0 Then
            If GB_RS1("PRO_CODE") = "G204" Or GB_RS1("PRO_CODE") = "G206" Or GB_RS1("PRO_CODE") = "G209" Or GB_RS1("PRO_CODE") = "G111" Or GB_RS1("PRO_CODE") = "G268" Or GB_RS1("PRO_CODE") = "G265" Or GB_RS1("PRO_CODE") = "G266" Or GB_RS1("PRO_CODE") = "G256" Then
                Frame3.Visible = True
                q = True
                str = "SELECT * FROM DISLIB.dmoprismer WHERE reqno = ? "
                Set ViewGBQR = rdb.CreateQuery("", str)

                gk = InStr(1, cmb_reqno, "$")

                ViewGBQR(0) = req_chglen(Mid(Trim(cmb_reqno), 1, gk - 1))
                Set ViewGBRS = ViewGBQR.OpenResultset(1, 3)
                If ViewGBRS.RowCount > 0 Then
                txt_btest.Text = Trim(ViewGBRS("OFINDING"))
                End If
                   If ViewGBRS.RowCount > 0 Then
                txt_bresult.Text = Trim(ViewGBRS("REPORT"))
                End If
            End If
        End If
        Else
        
            set_visible "frm_others"
            getresult_other
        End If
                
            If SELDADDENDUMFLAG = False Then
                SELDADDENDUM_QR(0) = mod_reqno
                Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
                SELDADDENDUMFLAG = True
            Else
                SELDADDENDUM_QR(0) = mod_reqno
                SELDADDENDUM_RS.MoreResults
                SELDADDENDUM_RS.Requery
            End If
            
            If SELDADDENDUM_RS.RowCount > 0 Then
                While Not SELDADDENDUM_RS.EOF
                    ADDENDUM = " ADDENDUM NO " & SELDADDENDUM_RS("ADDEND_NO") & Chr(13) & SELDADDENDUM_RS("ADD_MAX_LN") & Chr(13) & SELDADDENDUM_RS("ADDEND_DT") & " " & SELDADDENDUM_RS("ADDEND_TM") & " " & SELDADDENDUM_RS("USER_ID")
                    'cmb_addendumby.AddItem SELDADDENDUM_RS("USER_ID")
                    cmb_addendumby.Text = SELDADDENDUM_RS("USER_ID")
                SELDADDENDUM_RS.MoveNext
                Wend
                txt_addendum.Text = Trim(ADDENDUM)
                ADDENDUMFALG = True
            End If
        
      disable_field 'disable all fields
Exit Sub
MSG:
    MsgBox Err.Description
End Sub

'***** Shraddha 03/03/10 *****
Private Sub txt_test_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txt_testresult.SetFocus
End If
End Sub



Private Sub txt_testresult_KeyPress(KeyAscii As Integer)
'*********** Shraddha Pandit 03/03/10 ***********
If KeyAscii = 13 Then
'If MSFlexGrid1.col <> 2 Then
'    text1.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
'    text1.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
'    text1.Visible = True
'    text1.Text = ""
'    text1.SetFocus
'End If
text1.Left = 4000 '2520
text1.Top = 1450








text1.SetFocus
End If
'***********************************************
End Sub

Private Sub txt_urinecomment_GOTFOCUS()
On Error GoTo MSG
    Set objWord = CreateObject("word.application")
    objWord.Visible = False
    objWord.WindowState = wdWindowStateMaximize
    Set objDoc = objWord.Documents.add
    objDoc.Content = txt_urinecomment.Text
    objWord.Visible = True
    Exit Sub
MSG:
    'objWord.Repeat
    If Not Err.Number = -2147023174 Then
        MsgBox "error" & Err.Number
    End If
End Sub
Private Sub txt_urinecomment_LOSTFOCUS()
On Error GoTo MSG
    txt_urinecomment.Text = objDoc.Content
    objDoc.Content = ""
    ' Modified By Pankaj Baggal 03/09/2004
    'objDoc.SaveAs "c:\windows\system\frozen.doc"
    objDoc.Close (wdDoNotSaveChanges)

    objWord.quit
    Exit Sub
MSG:
    If Not Err.Number = -2147023174 Then
        MsgBox "error" & Err.Number
    End If

End Sub



Private Sub txt_urinother_GOTFOCUS()

On Error GoTo MSG
    Set objWord = CreateObject("word.application")
    
    objWord.Visible = False
    objWord.WindowState = wdWindowStateMaximize
    Set objDoc = objWord.Documents.add
    objDoc.Content = txt_urinother.Text
    objWord.Visible = True
    Exit Sub
MSG:
    'objWord.Repeat
    If Not Err.Number = -2147023174 Then
        MsgBox "error" & Err.Number
    End If
End Sub

Private Sub txt_urinother_LOSTFOCUS()
On Error GoTo MSG
    txt_urinother.Text = objDoc.Content
    objDoc.Content = ""
    ' Modified By Pankaj Baggal 03/09/2004
    'objDoc.SaveAs "c:\windows\system\frozen.doc"
    objDoc.Close (wdDoNotSaveChanges)
    objWord.quit
    Exit Sub
MSG:
    If Not Err.Number = -2147023174 Then
        MsgBox "error" & Err.Number
    End If
End Sub
Public Sub printpage(moda As String, mskval As String, reqval As String)
Me.Hide
'If moda = "BA" Then
    'cmd_view_Click
    VIEW_FLAG = True
    new_flag = False
    callstatus = True
    If IsNumeric(Trim(mskval)) Then
        opt_staff.Value = True
    Else
        Opt_patient.Value = True
    End If
    txt_deptcode.Text = "GZZ"
    txt_modcode.Text = moda
    MaskEdBox1.Text = mskval
    cmb_modality_Click
    
    
    'GZZREQPATH_QR(0) = req_rs("REQNNO")
    
    GZZREQPATH_QR(0) = req_chglen(req_sta)
    Set GZZREQPATH_RS = GZZREQPATH_QR.OpenResultset(1, 3)
    
    cmb_reqno.AddItem req_chglen(reqval) + "$" + GZZREQPATH_RS("pathno")
    
    cmb_reqno.Text = req_chglen(reqval) + "$" + GZZREQPATH_RS("pathno")
    cmb_reqno_Click
    cmd_print_Click
'End If

End Sub


Public Sub clearFrmSerology()
txtSerTestPerform.Text = ""
cmbSerRange.Text = ""
cmbSerResult.Text = ""
cmbProcalci.Text = ""
cmbProCalCon.Text = ""
txtInBold.Text = ""
txtInterpretation.Text = ""
txtMethod.Text = ""
txtLimitation.Text = ""
End Sub

Private Sub txtInterpretation_GotFocus()
On Error GoTo m14:
w.Visible = False
Set d = w.ActiveDocument
d.Content.Text = ""
d.Content.Text = txtInterpretation
w.Caption = "Full Report"
w.WindowState = wdWindowStateMaximize
d.Content.Text = txtInterpretation.Text
w.Visible = True
w.Activate
m14:
On Error Resume Next:
Set w = CreateObject("word.Application")
w.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
w.Caption = "Full Report"
'w.WindowState = wdWindowStateNormal
w.WindowState = wdWindowStateMaximize
Set d = w.ActiveDocument
d.Content.Text = txtInterpretation.Text
w.Visible = True
w.Activate
End Sub

Private Sub txtInterpretation_LostFocus()
On Error GoTo m15:
Set d = w.ActiveDocument
txtInterpretation.Text = d.Content.Text
'flagimp = False
d.Content.Text = ""
' Modified By Pankaj Baggal 03/09/2004
'd.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
d.Close (wdDoNotSaveChanges)
w.quit
'W1.Visible = False
Exit Sub
m15:
MsgBox Err.Description
End Sub

Private Sub txtLimitation_GotFocus()
On Error GoTo m14:
w.Visible = False
Set d = w.ActiveDocument
d.Content.Text = ""
d.Content.Text = txtLimitation
w.Caption = "Full Report"
w.WindowState = wdWindowStateMaximize
d.Content.Text = txtLimitation.Text
w.Visible = True
w.Activate
m14:
On Error Resume Next:
Set w = CreateObject("word.Application")
w.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
w.Caption = "Full Report"
'w.WindowState = wdWindowStateNormal
w.WindowState = wdWindowStateMaximize
Set d = w.ActiveDocument
d.Content.Text = txtLimitation.Text
w.Visible = True
w.Activate
End Sub

Private Sub txtLimitation_LostFocus()
On Error GoTo m15:
Set d = w.ActiveDocument
txtLimitation.Text = d.Content.Text
'flagimp = False
d.Content.Text = ""
' Modified By Pankaj Baggal 03/09/2004
'd.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
d.Close (wdDoNotSaveChanges)
w.quit
'W1.Visible = False
Exit Sub
m15:
MsgBox Err.Description
End Sub

Private Sub txtMethod_GotFocus()
On Error GoTo m14:
w.Visible = False
Set d = w.ActiveDocument
d.Content.Text = ""
d.Content.Text = txtMethod
w.Caption = "Full Report"
w.WindowState = wdWindowStateMaximize
d.Content.Text = txtMethod.Text
w.Visible = True
w.Activate
m14:
On Error Resume Next:
Set w = CreateObject("word.Application")
w.Documents.add ("c:\DISREPORTSEC\REPORT.dot")
w.Caption = "Full Report"
'w.WindowState = wdWindowStateNormal
w.WindowState = wdWindowStateMaximize
Set d = w.ActiveDocument
d.Content.Text = txtMethod.Text
w.Visible = True
w.Activate
End Sub

Private Sub txtMethod_LostFocus()
On Error GoTo m15:
Set d = w.ActiveDocument
txtMethod.Text = d.Content.Text
'flagimp = False
d.Content.Text = ""
' Modified By Pankaj Baggal 03/09/2004
'd.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
d.Close (wdDoNotSaveChanges)
w.quit
'W1.Visible = False
Exit Sub
m15:
MsgBox Err.Description
End Sub

