VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{9B8DFCFF-6726-11D6-99D4-000244193CE4}#1.0#0"; "Project1.ocx"
Begin VB.Form dis_inf 
   Caption         =   "Dis Common Screen"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
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
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Project1.UserControl1 UserControl11 
      Height          =   495
      Left            =   2160
      TabIndex        =   58
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   873
   End
   Begin VB.Frame common 
      Height          =   3855
      Left            =   0
      TabIndex        =   15
      Top             =   840
      Width           =   11895
      Begin VB.ComboBox cmb_prodis 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2400
         Sorted          =   -1  'True
         TabIndex        =   57
         Top             =   2520
         Width           =   3135
      End
      Begin VB.TextBox txt_durill 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10320
         TabIndex        =   31
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txt_clinhist 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         TabIndex        =   30
         Top             =   2040
         Width           =   5895
      End
      Begin VB.TextBox txt_reftodoc 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6840
         TabIndex        =   29
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txt_REFDOC 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         TabIndex        =   28
         Top             =   1680
         Width           =   2175
      End
      Begin VB.ComboBox cmb_modality 
         ForeColor       =   &H00FF0000&
         Height          =   315
         ItemData        =   "dis_inf.frx":0000
         Left            =   2400
         List            =   "dis_inf.frx":000D
         TabIndex        =   27
         Top             =   3480
         Width           =   3135
      End
      Begin VB.TextBox txt_modcode 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6840
         TabIndex        =   26
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox txt_bedno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10320
         TabIndex        =   25
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txt_ward 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6840
         TabIndex        =   24
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txt_pridocser 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         TabIndex        =   23
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txt_catgory 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6840
         TabIndex        =   22
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txt_status 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10320
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txt_Age 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   4680
         TabIndex        =   20
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txt_sex 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         TabIndex        =   19
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txt_name 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   4680
         TabIndex        =   18
         Top             =   240
         Width           =   6855
      End
      Begin VB.TextBox txt_dept 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6840
         TabIndex        =   17
         Top             =   2880
         Width           =   1335
      End
      Begin VB.ComboBox cmb_dept 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2400
         TabIndex        =   16
         Top             =   3000
         Width           =   3135
      End
      Begin MSMask.MaskEdBox Msk_caseno 
         Height          =   375
         Left            =   2400
         TabIndex        =   32
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   16711680
         HideSelection   =   0   'False
         MaxLength       =   8
         Mask            =   "??>/#####"
         PromptChar      =   "_"
      End
      Begin VB.Label lbl_provdis 
         AutoSize        =   -1  'True
         Caption         =   "Provisional Diagnosis"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   480
         TabIndex        =   50
         Top             =   2520
         Width           =   1830
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Duration of Illness"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   8595
         TabIndex        =   49
         Top             =   2040
         Width           =   1560
      End
      Begin VB.Label lbl_clinhist 
         Caption         =   "Relevant Clinical History"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   135
         TabIndex        =   48
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label lbl_reftodoc 
         AutoSize        =   -1  'True
         Caption         =   "Referred to Doctor"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5160
         TabIndex        =   47
         Top             =   1680
         Width           =   1605
      End
      Begin VB.Label lbl_refdoc 
         AutoSize        =   -1  'True
         Caption         =   "Referring Doctor/Service"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   150
         TabIndex        =   46
         Top             =   1680
         Width           =   2160
      End
      Begin VB.Label lbl_modality 
         AutoSize        =   -1  'True
         Caption         =   "Modality"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1590
         TabIndex        =   45
         Top             =   3480
         Width           =   720
      End
      Begin VB.Label lbl_modcode 
         AutoSize        =   -1  'True
         Caption         =   "Modality Code"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5580
         TabIndex        =   44
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label lbl_bedno 
         AutoSize        =   -1  'True
         Caption         =   "Bed No"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   9510
         TabIndex        =   43
         Top             =   1080
         Width           =   645
      End
      Begin VB.Label lbl_ward 
         AutoSize        =   -1  'True
         Caption         =   "Ward"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   6270
         TabIndex        =   42
         Top             =   1200
         Width           =   465
      End
      Begin VB.Label lbl_catgory 
         AutoSize        =   -1  'True
         Caption         =   "Category"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5970
         TabIndex        =   41
         Top             =   720
         Width           =   765
      End
      Begin VB.Label lbl_status 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   9600
         TabIndex        =   40
         Top             =   720
         Width           =   555
      End
      Begin VB.Label lbl_age 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4200
         TabIndex        =   39
         Top             =   720
         Width           =   345
      End
      Begin VB.Label lbl_pridocser 
         AutoSize        =   -1  'True
         Caption         =   "Primary Doctor/Service"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   315
         TabIndex        =   38
         Top             =   1200
         Width           =   1995
      End
      Begin VB.Label lbl_sex 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1980
         TabIndex        =   37
         Top             =   720
         Width           =   330
      End
      Begin VB.Label lbl_name 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4080
         TabIndex        =   36
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lbl_caseno 
         AutoSize        =   -1  'True
         Caption         =   "Case No."
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1515
         TabIndex        =   35
         Top             =   240
         Width           =   795
      End
      Begin VB.Label lbl_deptcode 
         AutoSize        =   -1  'True
         Caption         =   "Dept. Code"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5760
         TabIndex        =   34
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label lbl_dept 
         AutoSize        =   -1  'True
         Caption         =   "Requistion Placed to"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   525
         TabIndex        =   33
         Top             =   3000
         Width           =   1785
      End
   End
   Begin VB.TextBox txt_date 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   10560
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txt_time 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   10560
      TabIndex        =   10
      Top             =   480
      Width           =   1215
   End
   Begin VB.Frame buttons 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   7800
      Width           =   11775
      Begin VB.CommandButton cmd_new 
         Caption         =   "&New"
         Height          =   375
         Left            =   120
         MouseIcon       =   "dis_inf.frx":001D
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_view 
         Caption         =   "&View"
         Height          =   375
         Left            =   1560
         MouseIcon       =   "dis_inf.frx":0327
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_modify 
         Caption         =   "&Modify"
         Height          =   375
         Left            =   3000
         MouseIcon       =   "dis_inf.frx":0631
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_save 
         Caption         =   "&Save"
         Height          =   375
         Left            =   4560
         MouseIcon       =   "dis_inf.frx":093B
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   6120
         MouseIcon       =   "dis_inf.frx":0C45
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "&Print"
         Height          =   375
         Left            =   7440
         MouseIcon       =   "dis_inf.frx":0F4F
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_back 
         Caption         =   "&Back"
         Height          =   375
         Left            =   9000
         MouseIcon       =   "dis_inf.frx":1259
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_repeat 
         Caption         =   "&Repeat"
         Height          =   375
         Left            =   10440
         MouseIcon       =   "dis_inf.frx":1563
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame sp 
      ForeColor       =   &H000000FF&
      Height          =   2535
      Left            =   0
      TabIndex        =   9
      Top             =   4800
      Width           =   11895
      Begin VB.ListBox list_procedure 
         ForeColor       =   &H00FF0000&
         Height          =   1425
         Left            =   2400
         TabIndex        =   55
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox cmb_procedurecode 
         ForeColor       =   &H00FF0000&
         Height          =   315
         ItemData        =   "dis_inf.frx":186D
         Left            =   6840
         List            =   "dis_inf.frx":186F
         TabIndex        =   54
         Top             =   360
         Width           =   1695
      End
      Begin VB.ComboBox cmb_procedure 
         ForeColor       =   &H00FF0000&
         Height          =   315
         ItemData        =   "dis_inf.frx":1871
         Left            =   2400
         List            =   "dis_inf.frx":187E
         MouseIcon       =   "dis_inf.frx":1891
         MousePointer    =   99  'Custom
         TabIndex        =   52
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lbl_mprocedure 
         AutoSize        =   -1  'True
         Caption         =   "Multiple Procedure"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   600
         TabIndex        =   56
         Top             =   1200
         Width           =   1605
      End
      Begin VB.Label lbl_procedurecode 
         AutoSize        =   -1  'True
         Caption         =   "Procedure Code"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5400
         TabIndex        =   53
         Top             =   360
         Width           =   1620
      End
      Begin VB.Label lbl_procedure 
         Caption         =   "Name of Procedure"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   600
         TabIndex        =   51
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Label lbl_req 
      AutoSize        =   -1  'True
      Caption         =   "COMMON SCREEN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   3480
      TabIndex        =   14
      Top             =   480
      Width           =   2670
   End
   Begin VB.Label lbl_date 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   9960
      TabIndex        =   13
      Top             =   120
      Width           =   540
   End
   Begin VB.Label lbl_time 
      AutoSize        =   -1  'True
      Caption         =   "Time"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   9960
      TabIndex        =   12
      Top             =   480
      Width           =   540
   End
End
Attribute VB_Name = "dis_inf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmb_procedure_Click()
list_procedure.AddItem cmb_procedure.Text
End Sub

Private Sub cmd_back_Click()
DIS_MAIN.Show
Unload Me
End Sub


Private Sub Form_Unload(cancel As Integer)
Dim RS As rdoResultset
Dim RQ As rdoQuery
'For Each RS In rdb.rdoResultsets
'    RS.Close
'Next

'For Each RQ In rdb.rdoQueries
'    RQ.Close
'Next

End Sub

