VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form EXAMCOMP 
   Caption         =   "Exam  Complete Screen"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13185
   LinkTopic       =   "Form1"
   ScaleHeight     =   9240
   ScaleWidth      =   13185
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   480
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox txt_time 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   10320
      TabIndex        =   20
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txt_date 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   10320
      TabIndex        =   19
      Top             =   480
      Width           =   1215
   End
   Begin VB.Frame buttons 
      Height          =   735
      Left            =   0
      TabIndex        =   12
      Top             =   7800
      Width           =   11775
      Begin VB.CommandButton cmd_new 
         Caption         =   "&New"
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
         Left            =   720
         MousePointer    =   99  'Custom
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_view 
         Caption         =   "&View"
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
         Left            =   2160
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_modify 
         Caption         =   "&Modify"
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
         Left            =   3720
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_save 
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
         Left            =   5160
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "&Cancel"
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
         Left            =   6600
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_print 
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
         Left            =   8040
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_back 
         Caption         =   "&Back"
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
         Left            =   9600
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5055
      Left            =   0
      TabIndex        =   11
      Top             =   2760
      Width           =   11775
      Begin VB.TextBox TXT_LABNO 
         BackColor       =   &H00C0C0FF&
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
         Height          =   285
         Left            =   8520
         TabIndex        =   33
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox chk_collected 
         BackColor       =   &H8000000E&
         Caption         =   "Pro. Completed"
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
         Height          =   255
         Left            =   8520
         TabIndex        =   6
         Top             =   720
         Width           =   1815
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3615
         Left            =   1320
         TabIndex        =   5
         Top             =   360
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   6376
         _Version        =   393216
         Rows            =   3
         Cols            =   5
         FixedCols       =   0
         BackColor       =   16777215
         ForeColor       =   0
         ForeColorFixed  =   192
         AllowUserResizing=   3
         FormatString    =   "S.No |^    Requisation No.         |^    Case No.        |^                Procedure Name   |^           Exam Completed"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   0
      TabIndex        =   8
      Top             =   1200
      Width           =   11775
      Begin VB.ComboBox cmb_modality 
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
         Height          =   315
         Left            =   5880
         TabIndex        =   31
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txt_name 
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6720
         TabIndex        =   30
         Top             =   1080
         Width           =   3975
      End
      Begin MSMask.MaskEdBox msk_caseno 
         Height          =   375
         Left            =   4560
         TabIndex        =   4
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Frame Frame3 
         Height          =   855
         Left            =   720
         TabIndex        =   26
         Top             =   600
         Width           =   1335
         Begin VB.OptionButton opt_staff 
            Caption         =   "Staff"
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
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton Opt_patient 
            Caption         =   "Patient"
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
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.OptionButton OPT_CASENO 
         Caption         =   "Case No."
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
         Height          =   255
         Left            =   2640
         TabIndex        =   25
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton OPT_ALL 
         Caption         =   "All"
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
         Height          =   255
         Left            =   2640
         TabIndex        =   24
         Top             =   720
         Width           =   735
      End
      Begin VB.ComboBox cmb_location 
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
         Height          =   315
         Left            =   2280
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin MSMask.MaskEdBox msk_todt 
         Height          =   375
         Left            =   9360
         TabIndex        =   3
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   0
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
      Begin MSMask.MaskEdBox msk_fromdt 
         Height          =   375
         Left            =   4560
         TabIndex        =   2
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Modality"
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
         Height          =   195
         Left            =   5040
         TabIndex        =   32
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl_name 
         AutoSize        =   -1  'True
         Caption         =   "Name"
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
         Height          =   195
         Left            =   6120
         TabIndex        =   29
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lbl_location 
         AutoSize        =   -1  'True
         Caption         =   "Sample Collected at"
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
         Height          =   195
         Left            =   480
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   1710
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
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
         Height          =   195
         Left            =   3600
         TabIndex        =   10
         Top             =   720
         Width           =   885
      End
      Begin VB.Label lbl_fromdt 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
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
         Height          =   195
         Left            =   8520
         TabIndex        =   9
         Top             =   720
         Width           =   705
      End
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
      Left            =   1320
      TabIndex        =   34
      Top             =   0
      Width           =   8895
   End
   Begin VB.Label lbl_time 
      AutoSize        =   -1  'True
      Caption         =   "Time"
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
      Height          =   195
      Left            =   9720
      TabIndex        =   22
      Top             =   840
      Width           =   420
   End
   Begin VB.Label lbl_date 
      AutoSize        =   -1  'True
      Caption         =   "Date"
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
      Height          =   195
      Left            =   9720
      TabIndex        =   21
      Top             =   480
      Width           =   420
   End
   Begin VB.Label lbl_req 
      AutoSize        =   -1  'True
      Caption         =   "EXAM COMPLETION"
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
      Left            =   4440
      TabIndex        =   7
      Top             =   720
      Width           =   2925
   End
End
Attribute VB_Name = "EXAMCOMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim rlocation_qr As rdoQuery
Dim rlocation_rs As rdoResultset
Dim rlocationflag As Boolean
Dim COLLECTIONFLAG As Boolean
Dim SELRREQUISITN_QR As rdoQuery
Dim SELRREQUISITN_RS As rdoResultset
Dim SELRREQUISITNFLAG As Boolean
Dim updrrequisitn_qr As rdoQuery
Dim SELMA101_QR As rdoQuery
Dim SELMA101_RS As rdoResultset
Dim SELMA101FLAG As Boolean
Dim NEWFLAG As Boolean
Dim viewflag As Boolean
Dim allflag As Boolean
Dim casenoflag As Boolean
Dim SELREQNO_QR As rdoQuery
Dim SELREQNO_RS As rdoResultset
Dim SELREQNOFLAG As Boolean
Dim selta101_qr1 As rdoQuery
Dim selta101_rs1 As rdoResultset
Dim selta101flag1 As Boolean
Dim SELDREQSTAFFD_QR As rdoQuery
Dim SELDREQSTAFFD_RS As rdoResultset
Dim SELDREQSTAFFDFLAG As Boolean
Dim SELRMODLAITYQR As rdoQuery
Dim SELRMODLAITYRS As rdoResultset
Dim SELRMODLAITYFLAG As Boolean
Dim MODCODE As String
Dim UPDDREQUISITNQR As rdoQuery

Dim SELDREQUISITNQR As rdoQuery
Dim SELDREQUISITNRS As rdoResultset
Dim SELDREQUISITNFLAG As Boolean



Private Sub chk_collected_Click()
If chk_collected.Value = 1 Then
    MSFlexGrid1.Text = "Pro.Completed"
Else
    MSFlexGrid1.Text = ""
End If
End Sub



Private Sub chk_collected_KeyPress(KeyAscii As Integer)
On Error GoTo MSG:
If KeyAscii = 13 Then
If MSFlexGrid1.row = MSFlexGrid1.Rows - 1 Then
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = MSFlexGrid1.Rows - 1
End If
If MSFlexGrid1.row < MSFlexGrid1.Rows - 1 Then

        MSFlexGrid1.row = MSFlexGrid1.row + 1
        chk_collected = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4))
        chk_collected.SetFocus
        chk_collected.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
        
End If
End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub



Private Sub cmb_location_Click()
'msk_fromdt.Text = DATE11
'msk_todt.Text = DATE11
'msk_fromdt.SetFocus
''If allflag = True Then
''        OPT_ALL.Value = True
''    End If
''    If casenoflag = True Then
''        OPT_CASENO.Value = True
''        Msk_caseno.Enabled = True
''    End If
End Sub





Private Sub cmb_modality_Click()
i = InStr(cmb_modality.Text, "$")
MODCODE = Mid(cmb_modality, i + 1, 2)
If Trim(MODCODE) = "ED" Then
    chk_collected.Visible = False
    'TXT_LABNO.Visible = True
    MSFlexGrid1.FormatString = "S.No |^    Requisation No.         |^    Case No.        |^                Procedure Name   |^           ENDOSCOPY NO."
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.row = 1
    MSFlexGrid1.col = 4
    TXT_LABNO.Width = MSFlexGrid1.CellWidth
    TXT_LABNO.Height = MSFlexGrid1.CellHeight
    TXT_LABNO.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
    TXT_LABNO.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top



Else
   ' chk_collected.Visible = True
    'TXT_LABNO.Visible = False
    MSFlexGrid1.FormatString = "S.No |^    Requisation No.         |^    Case No.        |^                Procedure Name   |^           EXAM COMP."
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.row = 1
    MSFlexGrid1.col = 4
    chk_collected.Width = MSFlexGrid1.CellWidth
    chk_collected.Height = MSFlexGrid1.CellHeight
    chk_collected.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
    chk_collected.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top


End If
End Sub

Private Sub cmd_back_Click()
    cmb_location.Enabled = False
    msk_fromdt.Enabled = False
    msk_todt.Enabled = False
    chk_collected.Visible = False
    msk_fromdt.Mask = ""
    msk_fromdt.Text = ""
    msk_fromdt.Mask = "##/##/####"
    msk_todt.Mask = ""
    msk_todt.Text = ""
    msk_todt.Mask = "##/##/####"
    txt_name.Text = ""
    OPT_ALL.Value = False
    OPT_CASENO.Value = False
    msk_caseno.Enabled = False
    Opt_patient.Enabled = False
    opt_staff.Enabled = False
    MSFlexGrid1.Rows = 1
    TXT_LABNO.Text = ""
    cmb_modality.Text = ""
    msk_caseno.Text = ""
    chk_collected.Value = 0
btnenabl ("INI")
Me.MousePointer = 11
DIS_MAIN.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub cmd_cancel_Click()
MSFlexGrid1.Rows = 1
If OPT_ALL.Enabled = True Then
    cmb_location.Enabled = False
    msk_fromdt.Enabled = False
    msk_todt.Enabled = False
    chk_collected.Value = 0
    chk_collected.Visible = False
    msk_fromdt.Mask = ""
    msk_fromdt.Text = ""
    msk_fromdt.Mask = "##/##/####"
    msk_todt.Mask = ""
    msk_todt.Text = ""
    msk_todt.Mask = "##/##/####"
    txt_name.Text = ""
    OPT_ALL.Value = False
    OPT_CASENO.Value = False
    msk_caseno.Enabled = False
    Opt_patient.Enabled = False
    opt_staff.Enabled = False
    cmb_modality.Enabled = False
    TXT_LABNO.Visible = False
    allflag = True
End If
If OPT_CASENO.Enabled = True Then
    cmb_location.Enabled = False
    msk_fromdt.Enabled = False
    msk_todt.Enabled = False
    chk_collected.Value = 0
    chk_collected.Visible = False
    msk_fromdt.Mask = ""
    msk_fromdt.Text = ""
    txt_name.Text = ""
    msk_fromdt.Mask = "##/##/####"
    msk_todt.Mask = ""
    msk_todt.Text = ""
    msk_todt.Mask = "##/##/####"
    OPT_ALL.Value = False
    OPT_CASENO.Value = False
    msk_caseno.Enabled = False
    Opt_patient.Enabled = False
    opt_staff.Enabled = False
    msk_caseno.Mask = ""
    msk_caseno.Text = ""
    cmb_modality.Enabled = False
    msk_caseno.Enabled = False
    TXT_LABNO.Visible = False
    casenoflag = True
End If

btnenabl ("INI")
End Sub

Private Sub cmd_modify_Click()
btnenabl ("NEW")

i = InStr(cmb_modality.Text, "$")
MODCODE = Mid(cmb_modality, i + 1, 2)
If Trim(MODCODE) = "ED" Then
    chk_collected.Visible = False
    TXT_LABNO.Visible = True
    MSFlexGrid1.FormatString = "S.No |^    Requisation No.         |^    Case No.        |^                Procedure Name   |^           LAB NO."
    TXT_LABNO.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4)
Else
    chk_collected.Visible = True
    chk_collected.Value = 0
    TXT_LABNO.Visible = False
    MSFlexGrid1.FormatString = "S.No |^    Requisation No.         |^    Case No.        |^                Procedure Name   |^           EXAM COMP."
End If
End Sub

Private Sub cmd_new_Click()
On Error GoTo MSG:
NEWFLAG = True
btnenabl ("NEW")
'    If rlocationflag = False Then
'        Set rlocation_rs = rlocation_qr.OpenResultset(1, 3)
'        rlocationflag = True
'
'   Else
'        rlocation_rs.MoreResults
'        rlocation_rs.Requery
'    End If
'    If rlocation_rs.RowCount > 0 Then
'        cmb_location.clear
'        While Not rlocation_rs.EOF
'            cmb_location.AddItem Trim(rlocation_rs("LOCATION"))
'            rlocation_rs.MoveNext
'        Wend
'    End If
If SELRMODLAITYFLAG = False Then
    Set SELRMODLAITYRS = SELRMODLAITYQR.OpenResultset(1, 3)
    SELRMODLAITYFLAG = True
Else
    SELRMODLAITYRS.MoreResults
    SELRMODLAITYRS.Requery
End If
  If SELRMODLAITYRS.RowCount > 0 Then
    cmb_modality.clear
    While Not SELRMODLAITYRS.EOF
        cmb_modality.AddItem Trim(SELRMODLAITYRS("mODLITYDES")) & "$" & Trim(SELRMODLAITYRS("MODALITYCD"))
        SELRMODLAITYRS.MoveNext
    Wend
  End If
    'cmb_location.Enabled = True
    txt_name.Text = ""
    msk_fromdt.Enabled = False
    msk_todt.Enabled = False
    OPT_ALL.Enabled = True
    OPT_CASENO.Enabled = True
    Opt_patient.Enabled = True
    opt_staff.Enabled = True
    msk_caseno.Enabled = False
    Opt_patient.Value = False
    opt_staff.Value = False
    cmb_modality.Enabled = True
    cmb_modality.SetFocus
    'cmb_location.SetFocus
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub cmd_print_Click()
Dim var_date_f
Dim var_date_t
Dim str As String
str = "C"
On Error GoTo MSG:
    str = ipadd & " " & Format(DATE11, "dd/mm/yyyy") & " " & time1
    If OPT_ALL.Value = True Then
        var_date_f = Year(msk_fromdt.FormattedText) & "," & month(msk_fromdt.FormattedText) & "," & Day(msk_fromdt.FormattedText)
        var_date_t = Year(msk_todt.FormattedText) & "," & month(msk_todt.FormattedText) & "," & Day(msk_todt.FormattedText)
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\COLLECTION.RPT"
        CrystalReport1.SelectionFormula = "{RREQUISITN.REQNDT} >= Date(" & var_date_f & ") AND {RREQUISITN.REQNDT} <= Date(" & var_date_t & ")"
        CrystalReport1.Formulas(0) = "FROMDT = '" & msk_fromdt.FormattedText & "'"
        CrystalReport1.Formulas(1) = "TODT = '" & msk_todt.FormattedText & "'"
        CrystalReport1.Formulas(2) = "LOCATION = '" & cmb_location.Text & "'"
        CrystalReport1.Formulas(3) = "USRID = '" & str & "'"
        CrystalReport1.Formulas(4) = "EXAM = 'EXAM. COMPLETION'"
     End If
     If OPT_CASENO.Value = True Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\COLLECTIONC.RPT"
        CrystalReport1.SelectionFormula = "{RREQUISITN.CASENO}= '" & msk_caseno.FormattedText & "'"
        CrystalReport1.Formulas(0) = "LOCATION = '" & cmb_location.Text & "'"
        CrystalReport1.Formulas(1) = "USRID = '" & str & "'"
        CrystalReport1.Formulas(2) = "EXAM = 'EXAM. COMPLETION'"
     End If
     CrystalReport1.WindowState = crptMaximized
     CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
     CrystalReport1.Action = 1
btnenabl ("PRINT")
Exit Sub
MSG:
 MsgBox Err.Description
End Sub

Private Sub cmd_save_Click()
 On Error GoTo MSG:
 rdb.BeginTrans
 If Trim(MODCODE) = "ED" Then
    For i = 1 To MSFlexGrid1.Rows - 1
         MSFlexGrid1.col = 4
        If MSFlexGrid1.TextMatrix(i, 4) <> "" Then
           
            updrrequisitn_qr(0) = "E"
            updrrequisitn_qr(1) = GET_REVDATE(txt_date.Text)
            updrrequisitn_qr(2) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
            updrrequisitn_qr.Execute
            
            UPDDREQUISITNQR(0) = Trim(MSFlexGrid1.TextMatrix(i, 4))
            UPDDREQUISITNQR(1) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
            UPDDREQUISITNQR.Execute
            'SHWETA UPDATE RTIME LOG 02-04-2013
            If UPDRTIMELOG(Trim(req_chglen(MSFlexGrid1.TextMatrix(i, 1))), "EXAM") = True Then
            Else
            GoTo MSG
            End If
            'TILL HERE
        End If
    Next i
 End If
 If Trim(MODCODE) = "ST" Or Trim(MODCODE) = "2D" Or Trim(MODCODE) = "EK" Or Trim(MODCODE) = "PF" Then
    
    For i = 1 To MSFlexGrid1.Rows - 1
         MSFlexGrid1.col = 4
        If MSFlexGrid1.TextMatrix(i, 4) <> "" Then
           
            updrrequisitn_qr(0) = "E"
            updrrequisitn_qr(1) = GET_REVDATE(txt_date.Text)
            updrrequisitn_qr(2) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
            updrrequisitn_qr.Execute
            
            ''SHWETA UPDATE RTIME LOG 02-04-2013
            If UPDRTIMELOG(Trim(req_chglen(MSFlexGrid1.TextMatrix(i, 1))), "EXAM") = True Then
                   Else
                   GoTo MSG
                   End If
                   '''TILL HERE
               End If
               
    Next i
    
 End If
rdb.CommitTrans
MsgBox "Record is Successfully Saved."
btnenabl ("SAVE")
Exit Sub
MSG:
    rdb.RollbackTrans
    MsgBox Err.Description
    
End Sub

Private Sub cmd_view_Click()
On Error GoTo MSG:
btnenabl ("VIEW")
viewflag = True
If rlocationflag = False Then
        Set rlocation_rs = rlocation_qr.OpenResultset(1, 3)
        rlocationflag = True
         
   Else
        rlocation_rs.MoreResults
        rlocation_rs.Requery
    End If
    If rlocation_rs.RowCount > 0 Then
        cmb_location.clear
        While Not rlocation_rs.EOF
            cmb_location.AddItem Trim(rlocation_rs("LOCATION"))
            rlocation_rs.MoveNext
        Wend
    End If
If SELRMODLAITYFLAG = False Then
    Set SELRMODLAITYRS = SELRMODLAITYQR.OpenResultset(1, 3)
    SELRMODLAITYFLAG = True
Else
    SELRMODLAITYRS.MoreResults
    SELRMODLAITYRS.Requery
End If
  If SELRMODLAITYRS.RowCount > 0 Then
    cmb_modality.clear
    While Not SELRMODLAITYRS.EOF
        cmb_modality.AddItem Trim(SELRMODLAITYRS("mODLITYDES")) & "$" & Trim(SELRMODLAITYRS("MODALITYCD"))
        SELRMODLAITYRS.MoveNext
    Wend
  End If
    txt_name.Text = ""
chk_collected.Visible = False
cmb_location.Enabled = True
msk_fromdt.Enabled = True
msk_todt.Enabled = True
OPT_ALL.Enabled = True
OPT_CASENO.Enabled = True
Opt_patient.Enabled = True
opt_staff.Enabled = True
Opt_patient.Value = False
opt_staff.Value = False
cmb_modality.Enabled = True
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo MSG:
NEWFLAG = False
viewflag = False
allflag = False
casenoflag = False
datetime
txt_date.Text = DATE11
txt_time.Text = time1
chk_collected.Visible = False
btnenabl ("INI")
rlocationflag = False

If EXAMCOMPFLAG = False Then
str = "select * from rislib.rlocation where dept_code='FZZ'"
Set rlocation_qr = rdb.CreateQuery("RLOCATION11_QREX", str)

'STR = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNDT BETWEEN ? AND ? AND ((REQNSTAT= ? OR REQNSTAT=? OR REQNSTAT=?)AND MEMODT IS NOT NULL  ) AND (DEPT_CODE='CAZ' OR DEPT_CODE='CBZ') AND MODALITYCD=? "
str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNDT BETWEEN ? AND ? AND ((REQNSTAT= ? OR REQNSTAT=? OR REQNSTAT=?)  ) AND (DEPT_CODE='CAZ' OR DEPT_CODE='CBZ') AND MODALITYCD=? "
'STR = "SELECT * FROM DISLIB.COLLECTVIEW WHERE REQNDT BETWEEN ? AND ?  "
Set SELRREQUISITN_QR = rdb.CreateQuery("SELRREQUISITN_QR12EX", str)

str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE=?"
Set SELMA101_QR = rdb.CreateQuery("SELMA101_QR11EX", str)


str = "UPDATE RISLIB.RREQUISITN SET  REQNSTAT=?,EXAMDT=? WHERE REQNNO=?"
Set updrrequisitn_qr = rdb.CreateQuery("UPDREQUISITN_QR1EX", str)

'STR = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND ((REQNSTAT=? OR REQNSTAT =? OR REQNSTAT =?) AND MEMODT IS NOT NULL)AND MODALITYCD=?"
str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND ((REQNSTAT=? OR REQNSTAT =? OR REQNSTAT =?) )AND MODALITYCD=?"
Set SELREQNO_QR = rdb.CreateQuery("SELREQNO_QR1EX", str)

str = "select * from patoolslib.ta101 where case_no=?"
Set selta101_qr1 = rdb.CreateQuery("selta101_qr1qEX", str)

str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE EMPCODE=?"
Set SELDREQSTAFFD_QR = rdb.CreateQuery("SELDREQSTAFFD_QR1EX", str)

str = "SELECT * FROM RISLIB.RMODALITY WHERE (DEPT_CODE='CAZ' OR DEPT_CODE='CBZ')"
Set SELRMODLAITYQR = rdb.CreateQuery("SELRMODLAITYQRCCEX", str)

str = "UPDATE DISLIB.DREQUISITN SET PATHNO=? WHERE REQNO=?"
Set UPDDREQUISITNQR = rdb.CreateQuery("UPDDREQUISITNQRED", str)

str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO=?"
Set SELDREQUISITNQR = rdb.CreateQuery("SELDREQUISITNQRED", str)

EXAMCOMPFLAG = True
End If
MSFlexGrid1.row = 1
MSFlexGrid1.col = 4
chk_collected.Width = MSFlexGrid1.CellWidth
chk_collected.Height = MSFlexGrid1.CellHeight
chk_collected.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
chk_collected.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
'TXT_FILMQTY.Width = MSF.CellWidth
'TXT_FILMQTY.Height = MSF.CellHeight
'TXT_FILMQTY.Left = MSF.CellLeft + MSF.Left
'TXT_FILMQTY.Top = MSF.CellTop + MSF.Top
cmb_location.Enabled = False
msk_fromdt.Enabled = False
msk_todt.Enabled = False
OPT_ALL.Enabled = False
OPT_CASENO.Enabled = False
Opt_patient.Enabled = False
opt_staff.Enabled = False
msk_caseno.Enabled = False
cmb_modality.Enabled = False
Exit Sub
MSG:
MsgBox Err.Description
End Sub
Public Sub btnenabl(cap As String)

If cap = "NEW" Then
    cmd_save.Enabled = True
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    CMD_NEW.Enabled = False
    cmd_modify.Enabled = False
    cmd_print.Enabled = False
    cmd_view.Enabled = False
ElseIf cap = "VIEW" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_modify.Enabled = True
    CMD_NEW.Enabled = False
    cmd_save.Enabled = False
    cmd_print.Enabled = True
    cmd_view.Enabled = False
ElseIf cap = "SAVE" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_modify.Enabled = False
    CMD_NEW.Enabled = True
    cmd_save.Enabled = False
    cmd_print.Enabled = True
    cmd_view.Enabled = True
ElseIf cap = "INI" Then
    cmd_cancel.Enabled = False
    cmd_back.Enabled = True
    cmd_modify.Enabled = False
    CMD_NEW.Enabled = True
    cmd_save.Enabled = False
    cmd_print.Enabled = False
    cmd_view.Enabled = True
ElseIf cap = "MODIFY" Then
    cmd_save.Enabled = True
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    CMD_NEW.Enabled = False
    cmd_modify.Enabled = False
    cmd_print.Enabled = False
    cmd_view.Enabled = False
End If


End Sub

Private Sub MSFlexGrid1_Click()
On Error GoTo MSG:
TXT_LABNO.Text = ""
If TXT_LABNO.Text <> "" Then
 MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = TXT_LABNO.Text
End If
If Trim(MODCODE) = "EK" Or Trim(MODCODE) = "2D" Or Trim(MODCODE) = "PF" Or Trim(MODCODE) = "ST" Then
 TXT_LABNO.Visible = False
 MSFlexGrid1.col = 4
'If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then

        'MSFlexGrid1.Row = MSFlexGrid1.Row + 1
        chk_collected = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4))
        chk_collected.SetFocus
        chk_collected.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
        If Trim(MSFlexGrid1.Text) = "COLLECTED" Then
        chk_collected.Value = 1
        Else
        chk_collected.Value = 0
        End If
'End If
End If
'*************** 0n 30/12/2004
    MSFlexGrid1.col = 4
 'If txt_pathno.Visible <> True Then
    TXT_LABNO.Width = MSFlexGrid1.CellWidth
    TXT_LABNO.Height = MSFlexGrid1.CellHeight
    TXT_LABNO.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
    TXT_LABNO.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
 'End If

Exit Sub
MSG:
MsgBox Err.Description
End Sub


Private Sub msk_caseno_LostFocus()
Dim i As Integer
On Error GoTo MSG:
If msk_caseno.Text = "" Then
    Exit Sub
End If
If Opt_patient.Value = True And OPT_CASENO.Value = True Then
    CASENO = padzero(5, "C", msk_caseno.Text, 3)
    msk_caseno.Text = UCase(CASENO)
    If selta101flag1 = False Then
        selta101_qr1(0) = msk_caseno.FormattedText
        Set selta101_rs1 = selta101_qr1.OpenResultset(1, 3)
        selta101flag1 = True
    Else
        selta101_qr1(0) = msk_caseno.FormattedText
        selta101_rs1.MoreResults
        selta101_rs1.Requery
    End If
    If selta101_rs1.RowCount > 0 Then
        txt_name = Trim(selta101_rs1("tag")) & " " & Trim(selta101_rs1("first_name")) & " " & Trim(selta101_rs1("mid_name")) & " " & Trim(selta101_rs1("last_name"))
    End If
If NEWFLAG = True Then
        NEWFLAG = False
        If SELREQNOFLAG = False Then
            SELREQNO_QR(0) = msk_caseno.FormattedText
            SELREQNO_QR(1) = "P"
            SELREQNO_QR(2) = "P"
            SELREQNO_QR(3) = "P"
            i = InStr(cmb_modality.Text, "$")
            SELREQNO_QR(4) = Mid(cmb_modality, i + 1, 2)
            Set SELREQNO_RS = SELREQNO_QR.OpenResultset(1, 3)
            SELREQNOFLAG = True
        Else
           SELREQNO_QR(0) = msk_caseno.FormattedText
           SELREQNO_QR(1) = "P"
           SELREQNO_QR(2) = "P"
           SELREQNO_QR(3) = "P"
           i = InStr(cmb_modality.Text, "$")
            SELREQNO_QR(4) = Mid(cmb_modality, i + 1, 2)
           SELREQNO_RS.MoreResults
           SELREQNO_RS.Requery
            
        End If

        If SELREQNO_RS.RowCount > 0 Then
             i = 1
                MSFlexGrid1.Rows = 1
                While Not SELREQNO_RS.EOF
                If SELMA101FLAG = False Then
                    SELMA101_QR(0) = SELREQNO_RS("SER_CODE")
                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                    SELMA101FLAG = True
                Else
                    SELMA101_QR(0) = SELREQNO_RS("SER_CODE")
                    SELMA101_RS.MoreResults
                    SELMA101_RS.Requery
               End If
                If SELMA101_RS.RowCount > 0 Then
                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELREQNO_RS("REQNNO")) & vbTab & SELREQNO_RS("CASENO") & vbTab & Trim(SELMA101_RS("SER_DET"))
                    i = i + 1
                End If
                SELREQNO_RS.MoveNext
                Wend
                If Trim(MODCODE) = "ED" Then
                    TXT_LABNO.Visible = True
                    chk_collected.Visible = False
                Else
                    TXT_LABNO.Visible = False
                    chk_collected.Visible = True
                End If
        End If
End If
If viewflag = True Then
             viewflag = FLASE
        
    If SELREQNOFLAG = False Then
        SELREQNO_QR(0) = msk_caseno.FormattedText
        SELREQNO_QR(1) = "E"
        SELREQNO_QR(2) = "E"
        SELREQNO_QR(3) = "E"
           i = InStr(cmb_modality.Text, "$")
            SELREQNO_QR(4) = Mid(cmb_modality, i + 1, 2)
        Set SELREQNO_RS = SELREQNO_QR.OpenResultset(1, 3)
        SELREQNOFLAG = True
    Else
        SELREQNO_QR(0) = msk_caseno.FormattedText
        SELREQNO_QR(1) = "E"
        SELREQNO_QR(2) = "E"
        SELREQNO_QR(3) = "E"
           i = InStr(cmb_modality.Text, "$")
            SELREQNO_QR(4) = Mid(cmb_modality, i + 1, 2)
        SELREQNO_RS.MoreResults
        SELREQNO_RS.Requery
    End If
          

    If SELREQNO_RS.RowCount > 0 Then
        i = 1
        MSFlexGrid1.Rows = 1
        While Not SELREQNO_RS.EOF
             If SELMA101FLAG = False Then
                 SELMA101_QR(0) = SELREQNO_RS("SER_CODE")
                 Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                 SELMA101FLAG = True
             Else
                 SELMA101_QR(0) = SELREQNO_RS("SER_CODE")
                 SELMA101_RS.MoreResults
                 SELMA101_RS.Requery
            End If
            
            If SELDREQUISITNFLAG = False Then
                SELDREQUISITNQR(0) = SELREQNO_RS("REQNNO")
                Set SELDREQUISITNRS = SELDREQUISITNQR.OpenResultset(1, 3)
                SELDREQUISITNFLAG = True
            Else
                 SELDREQUISITNQR(0) = SELREQNO_RS("REQNNO")
                 SELDREQUISITNRS.MoreResults
                 SELDREQUISITNRS.Requery
            End If
            
            If SELMA101_RS.RowCount > 0 Then
                If Trim(SELREQNO_RS("REQNSTAT")) = "E" Then
                    If Trim(MODCODE) = "ED" Then
                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELREQNO_RS("REQNNO")) & vbTab & SELREQNO_RS("CASENO") & vbTab & Trim(SELMA101_RS("SER_DET")) & vbTab & SELDREQUISITNRS("PATHNO")
                    Else
                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELREQNO_RS("REQNNO")) & vbTab & SELREQNO_RS("CASENO") & vbTab & Trim(SELMA101_RS("SER_DET")) & vbTab & "Pro.Completed"
                    End If
                Else
                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELREQNO_RS("REQNNO")) & vbTab & SELREQNO_RS("CASENO") & vbTab & Trim(SELMA101_RS("SER_DET"))
                End If
            i = i + 1
            End If
            SELREQNO_RS.MoveNext
        Wend
        chk_collected.Visible = False
    Else
        MsgBox "There are No Pending Requisition No."
    End If


End If
End If
If opt_staff.Value = True Then
    STAFF
End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub msk_fromdt_LostFocus()
On Error GoTo MSG:
If IsDate(GET_REVDATE(msk_fromdt.FormattedText)) = False Then
             MsgBox "Pleae ENTER the correct DATE."
                msk_fromdt.Mask = ""
                msk_fromdt.Text = ""
                msk_fromdt.Mask = "##/##/####"
'             MaskEd_from_date.Text = ""
             msk_fromdt.SetFocus
             'Exit Sub
             Else
             
             msk_todt.SetFocus
    End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub msk_todt_LostFocus()
'Dim i As Integer
On Error GoTo MSG:
If NEWFLAG = True Then
    If SELRREQUISITNFLAG = False Then
        SELRREQUISITN_QR(0) = msk_fromdt.FormattedText
        SELRREQUISITN_QR(1) = msk_todt.FormattedText
        SELRREQUISITN_QR(2) = "P"
        SELRREQUISITN_QR(3) = "P"
        SELRREQUISITN_QR(4) = "P"
        i = InStr(cmb_modality.Text, "$")
        SELRREQUISITN_QR(5) = Mid(cmb_modality, i + 1, 2)
        Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
    Else
        SELRREQUISITN_QR(0) = msk_fromdt.FormattedText
        SELRREQUISITN_QR(1) = msk_todt.FormattedText
        SELRREQUISITN_QR(2) = "P"
        SELRREQUISITN_QR(3) = "P"
        SELRREQUISITN_QR(4) = "P"
        i = InStr(cmb_modality.Text, "$")
        SELRREQUISITN_QR(5) = Mid(cmb_modality, i + 1, 2)
        SELRREQUISITN_RS.MoreResults
        SELRREQUISITN_RS.Requery
    End If

    NEWFLAG = False
    If SELRREQUISITN_RS.RowCount > 0 Then
        i = 1
        MSFlexGrid1.Rows = 1
        While Not SELRREQUISITN_RS.EOF
        If SELMA101FLAG = False Then
            SELMA101_QR(0) = SELRREQUISITN_RS("SER_CODE")
            Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
            SELMA101FLAG = True
        Else
            SELMA101_QR(0) = SELRREQUISITN_RS("SER_CODE")
            SELMA101_RS.MoreResults
            SELMA101_RS.Requery
       End If
        If SELMA101_RS.RowCount > 0 Then
        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("REQNNO")) & vbTab & SELRREQUISITN_RS("CASENO") & vbTab & Trim(SELMA101_RS("SER_DET"))
        i = i + 1
        End If
        SELRREQUISITN_RS.MoveNext
        Wend
        If Trim(MODCODE) = "ED" Then
            TXT_LABNO.Visible = True
            chk_collected.Visible = False
        Else
            TXT_LABNO.Visible = False
            chk_collected.Visible = True
        End If
    Else
        MsgBox "There are No Pending Requisition No."
    End If
End If
If viewflag = True Then
        viewflag = FLASE
        
        If SELRREQUISITNFLAG = False Then
        SELRREQUISITN_QR(0) = msk_fromdt.FormattedText
        SELRREQUISITN_QR(1) = msk_todt.FormattedText
        SELRREQUISITN_QR(2) = "E"
        SELRREQUISITN_QR(3) = "E"
         SELRREQUISITN_QR(4) = "E"
        i = InStr(cmb_modality.Text, "$")
        SELRREQUISITN_QR(5) = Mid(cmb_modality, i + 1, 2)
        Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
    Else
        SELRREQUISITN_QR(0) = msk_fromdt.FormattedText
        SELRREQUISITN_QR(1) = msk_todt.FormattedText
        SELRREQUISITN_QR(2) = "E"
        SELRREQUISITN_QR(3) = "E"
         SELRREQUISITN_QR(4) = "E"
        i = InStr(cmb_modality.Text, "$")
        SELRREQUISITN_QR(5) = Mid(cmb_modality, i + 1, 2)
        SELRREQUISITN_RS.MoreResults
        SELRREQUISITN_RS.Requery
    End If
        
        

        If SELRREQUISITN_RS.RowCount > 0 Then
        i = 1
        MSFlexGrid1.Rows = 1
        While Not SELRREQUISITN_RS.EOF
        If SELMA101FLAG = False Then
            SELMA101_QR(0) = SELRREQUISITN_RS("SER_CODE")
            Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
            SELMA101FLAG = True
        Else
            SELMA101_QR(0) = SELRREQUISITN_RS("SER_CODE")
            SELMA101_RS.MoreResults
            SELMA101_RS.Requery
       End If
       
            If SELDREQUISITNFLAG = False Then
                SELDREQUISITNQR(0) = SELRREQUISITN_RS("REQNNO")
                Set SELDREQUISITNRS = SELDREQUISITNQR.OpenResultset(1, 3)
                SELDREQUISITNFLAG = True
            Else
                 SELDREQUISITNQR(0) = SELRREQUISITN_RS("REQNNO")
                 SELDREQUISITNRS.MoreResults
                 SELDREQUISITNRS.Requery
            End If
       
       
        If SELMA101_RS.RowCount > 0 Then
        If Trim(SELRREQUISITN_RS("REQNSTAT")) = "E" Then
            If Trim(MODCODE) = "ED" Then
                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("REQNNO")) & vbTab & SELRREQUISITN_RS("CASENO") & vbTab & Trim(SELMA101_RS("SER_DET")) & vbTab & SELDREQUISITNRS("PATHNO")
            Else
                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("REQNNO")) & vbTab & SELRREQUISITN_RS("CASENO") & vbTab & Trim(SELMA101_RS("SER_DET")) & vbTab & "Pro.Completed."
            End If
        Else
            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("REQNNO")) & vbTab & SELRREQUISITN_RS("CASENO") & vbTab & Trim(SELMA101_RS("SER_DET"))
        End If
        i = i + 1
        End If
        SELRREQUISITN_RS.MoveNext
        Wend
        chk_collected.Visible = False
    Else
        MsgBox "There are No Pending Requisition No."
    End If


End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub Opt_all_Click()
    If OPT_ALL.Value = True Then
        msk_fromdt.Enabled = True
        msk_todt.Enabled = True
        Opt_patient.Enabled = True
        opt_staff.Enabled = True
        OPT_CASENO.Enabled = True
        msk_caseno.Enabled = False
        msk_fromdt.Text = DATE11
        msk_todt.Text = DATE11
        msk_fromdt.SetFocus
        casenoflag = False
        Opt_patient.Value = False
        opt_staff.Value = False
        lbl_name.Visible = False
        txt_name.Visible = False
    End If
    
End Sub

Private Sub OPT_CASENO_Click()
Opt_patient.Value = True
lbl_name.Visible = True
txt_name.Visible = True
End Sub

Private Sub Opt_patient_Click()
If Opt_patient.Value = True Then
    OPT_CASENO.Caption = "Case No."
    msk_caseno.Mask = "??>/#####"
    OPT_CASENO.Enabled = True
    OPT_ALL.Enabled = False
    msk_fromdt.Enabled = False
    msk_todt.Enabled = False
    msk_caseno.Enabled = True
    allflag = False
    msk_caseno.SetFocus
End If
End Sub

Private Sub opt_staff_Click()
If opt_staff.Value = True Then
    OPT_CASENO.Caption = "CC No."
    msk_caseno.Mask = "######"
    OPT_CASENO.Enabled = True
    'news
    OPT_CASENO.Value = True
    OPT_ALL.Enabled = False
    msk_fromdt.Enabled = False
    msk_todt.Enabled = False
    msk_caseno.Enabled = True
    msk_caseno.SetFocus
End If
End Sub

Public Sub STAFF()
    If SELDREQSTAFFDFLAG = False Then
        SELDREQSTAFFD_QR(0) = msk_caseno.FormattedText
        Set SELDREQSTAFFD_RS = SELDREQSTAFFD_QR.OpenResultset(1, 3)
        SELDREQSTAFFDFLAG = True
    Else
        SELDREQSTAFFD_QR(0) = msk_caseno.FormattedText
        SELDREQSTAFFD_RS.MoreResults
        SELDREQSTAFFD_RS.Requery
    End If
    If SELDREQSTAFFD_RS.RowCount > 0 Then
       SELDREQSTAFFD_RS.MoveLast
        txt_name = SELDREQSTAFFD_RS("PNAME")
    End If

If NEWFLAG = True Then
        NEWFLAG = False
        If SELREQNOFLAG = False Then
            SELREQNO_QR(0) = msk_caseno.FormattedText
            SELREQNO_QR(1) = "P"
            SELREQNO_QR(2) = "M"
            SELREQNO_QR(3) = "A"
            i = InStr(cmb_modality.Text, "$")
            SELREQNO_QR(4) = Mid(cmb_modality, i + 1, 2)
            Set SELREQNO_RS = SELREQNO_QR.OpenResultset(1, 3)
            SELREQNOFLAG = True
        Else
           SELREQNO_QR(0) = msk_caseno.FormattedText
           SELREQNO_QR(1) = "P"
           SELREQNO_QR(2) = "M"
           SELREQNO_QR(3) = "A"
           i = InStr(cmb_modality.Text, "$")
            SELREQNO_QR(4) = Mid(cmb_modality, i + 1, 2)
           SELREQNO_RS.MoreResults
           SELREQNO_RS.Requery
            
        End If

        If SELREQNO_RS.RowCount > 0 Then
             i = 1
                MSFlexGrid1.Rows = 1
                While Not SELREQNO_RS.EOF
                If SELMA101FLAG = False Then
                    SELMA101_QR(0) = SELREQNO_RS("SER_CODE")
                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                    SELMA101FLAG = True
                Else
                    SELMA101_QR(0) = SELREQNO_RS("SER_CODE")
                    SELMA101_RS.MoreResults
                    SELMA101_RS.Requery
               End If
                If SELMA101_RS.RowCount > 0 Then
                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELREQNO_RS("REQNNO")) & vbTab & SELREQNO_RS("CASENO") & vbTab & Trim(SELMA101_RS("SER_DET"))
                    i = i + 1
                End If
                SELREQNO_RS.MoveNext
                Wend
                If Trim(MODCODE) = "ED" Then
                    TXT_LABNO.Visible = True
                    chk_collected.Visible = False
                Else
                    TXT_LABNO.Visible = False
                    chk_collected.Visible = True
                End If
        End If
End If
If viewflag = True Then
             viewflag = FLASE
        
    If SELREQNOFLAG = False Then
        SELREQNO_QR(0) = msk_caseno.FormattedText
        SELREQNO_QR(1) = "E"
        SELREQNO_QR(2) = "E"
        SELREQNO_QR(3) = "E"
           i = InStr(cmb_modality.Text, "$")
            SELREQNO_QR(4) = Mid(cmb_modality, i + 1, 2)
        Set SELREQNO_RS = SELREQNO_QR.OpenResultset(1, 3)
        SELREQNOFLAG = True
    Else
        SELREQNO_QR(0) = msk_caseno.FormattedText
        SELREQNO_QR(1) = "E"
        SELREQNO_QR(2) = "E"
        SELREQNO_QR(3) = "E"
           i = InStr(cmb_modality.Text, "$")
            SELREQNO_QR(4) = Mid(cmb_modality, i + 1, 2)
        SELREQNO_RS.MoreResults
        SELREQNO_RS.Requery
    End If
          

    If SELREQNO_RS.RowCount > 0 Then
        i = 1
        MSFlexGrid1.Rows = 1
        While Not SELREQNO_RS.EOF
             If SELMA101FLAG = False Then
                 SELMA101_QR(0) = SELREQNO_RS("SER_CODE")
                 Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                 SELMA101FLAG = True
             Else
                 SELMA101_QR(0) = SELREQNO_RS("SER_CODE")
                 SELMA101_RS.MoreResults
                 SELMA101_RS.Requery
            End If
            
            If SELDREQUISITNFLAG = False Then
                SELDREQUISITNQR(0) = SELREQNO_RS("REQNNO")
                Set SELDREQUISITNRS = SELDREQUISITNQR.OpenResultset(1, 3)
                SELDREQUISITNFLAG = True
            Else
                 SELDREQUISITNQR(0) = SELREQNO_RS("REQNNO")
                 SELDREQUISITNRS.MoreResults
                 SELDREQUISITNRS.Requery
            End If
            
            If SELMA101_RS.RowCount > 0 Then
                If Trim(SELREQNO_RS("REQNSTAT")) = "E" Then
                    If Trim(MODCODE) = "ED" Then
                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELREQNO_RS("REQNNO")) & vbTab & SELREQNO_RS("CASENO") & vbTab & Trim(SELMA101_RS("SER_DET")) & vbTab & SELDREQUISITNRS("PATHNO")
                    Else
                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELREQNO_RS("REQNNO")) & vbTab & SELREQNO_RS("CASENO") & vbTab & Trim(SELMA101_RS("SER_DET")) & vbTab & "Pro.Completed"
                    End If
                Else
                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELREQNO_RS("REQNNO")) & vbTab & SELREQNO_RS("CASENO") & vbTab & Trim(SELMA101_RS("SER_DET"))
                End If
            i = i + 1
            End If
            SELREQNO_RS.MoveNext
        Wend
        chk_collected.Visible = False
    Else
        MsgBox "There are No Pending Requisition No."
    End If


End If



End Sub

Private Sub TXT_LABNO_Change()
If TXT_LABNO.Text <> "" Then
     MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = TXT_LABNO.Text
End If
End Sub

Private Sub TXT_LABNO_GotFocus()
TXT_LABNO.Text = ""
TXT_LABNO.BackColor = vbYellow
End Sub

Private Sub TXT_LABNO_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then
    If MSFlexGrid1.row <> 1 Then
    MSFlexGrid1.row = MSFlexGrid1.row - 1
    TXT_LABNO.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
    TXT_LABNO = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4))
    'MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = ""
    
    End If
End If
If KeyCode = 40 Then
    If MSFlexGrid1.row <> MSFlexGrid1.Rows - 1 Then
        MSFlexGrid1.row = MSFlexGrid1.row + 1
        TXT_LABNO.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
        TXT_LABNO = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4))
    End If
End If
End Sub

Private Sub TXT_LABNO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TXT_LABNO.Text) <> "" Then
        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = TXT_LABNO.Text
    End If
    i = MSFlexGrid1.Rows - 1
    If i = MSFlexGrid1.row Then
        cmd_save.SetFocus
        Exit Sub
    End If

    If MODIFYFLAG = False Then
        If MSFlexGrid1.row = MSFlexGrid1.Rows - 1 Then
                MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = MSFlexGrid1.Rows - 1
        End If
        If MSFlexGrid1.row < MSFlexGrid1.Rows - 1 Then
                MSFlexGrid1.row = MSFlexGrid1.row + 1
                TXT_LABNO = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4))
                'chk_rcvd = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5))
                TXT_LABNO.Visible = True
                'chk_rcvd.Visible = True
                'chk_rcvd.SetFocus
                TXT_LABNO.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                'chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                'MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = ""
        End If
    Else
        If MSFlexGrid1.row = MSFlexGrid1.Rows - 1 Then
                MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = MSFlexGrid1.Rows - 1
        End If
        If MSFlexGrid1.row < MSFlexGrid1.Rows - 1 Then
                MSFlexGrid1.row = MSFlexGrid1.row + 1
                TXT_LABNO = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4))
                'chk_rcvd = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5))
                TXT_LABNO.Visible = True
                'chk_rcvd.Visible = True
                'chk_rcvd.SetFocus
                TXT_LABNO.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                'chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                
                
        End If
    End If
End If
End Sub

Private Sub TXT_LABNO_LostFocus()
If TXT_LABNO.Text <> "" Then
     MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = TXT_LABNO.Text
End If
End Sub
