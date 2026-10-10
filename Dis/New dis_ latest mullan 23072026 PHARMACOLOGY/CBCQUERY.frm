VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form CBCQUERY 
   Caption         =   "CBC QUERY"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10665
   LinkTopic       =   "Form2"
   ScaleHeight     =   8820
   ScaleWidth      =   10665
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   0
      TabIndex        =   11
      Top             =   600
      Width           =   10575
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   840
         TabIndex        =   17
         Top             =   840
         Width           =   1335
         Begin VB.OptionButton Opt_patient 
            BackColor       =   &H00E0E0E0&
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
            TabIndex        =   19
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton opt_staff 
            BackColor       =   &H00E0E0E0&
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
            TabIndex        =   18
            Top             =   600
            Width           =   1095
         End
         Begin VB.Shape Shape1 
            BorderWidth     =   2
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   735
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.ComboBox cmb_dept 
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
         Left            =   2400
         TabIndex        =   16
         Top             =   360
         Width           =   2415
      End
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
         Left            =   6600
         TabIndex        =   15
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txt_name 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3240
         TabIndex        =   14
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox txt_age 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7320
         TabIndex        =   13
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txt_sex 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8760
         TabIndex        =   12
         Top             =   1440
         Width           =   855
      End
      Begin MSMask.MaskEdBox msk_caseno 
         Height          =   375
         Left            =   3240
         TabIndex        =   20
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Case No"
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
         Left            =   2400
         TabIndex        =   26
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
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
         Left            =   960
         TabIndex        =   25
         Top             =   360
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   5520
         TabIndex        =   24
         Top             =   360
         Width           =   840
      End
      Begin VB.Label lbl_name 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   2400
         TabIndex        =   23
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lbl_age 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
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
         Left            =   6840
         TabIndex        =   22
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label lbl_sex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
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
         Left            =   8280
         TabIndex        =   21
         Top             =   1440
         Width           =   330
      End
   End
   Begin VB.Frame Frame4 
      Height          =   4935
      Left            =   0
      TabIndex        =   9
      Top             =   3840
      Width           =   10575
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4335
         Left            =   360
         TabIndex        =   10
         Top             =   480
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   7646
         _Version        =   393216
         FixedRows       =   0
      End
   End
   Begin VB.Frame Frame3 
      Enabled         =   0   'False
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   2520
      Width           =   10575
      Begin VB.TextBox txt_num 
         Height          =   285
         Left            =   3360
         TabIndex        =   8
         Top             =   720
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
         Begin MSMask.MaskEdBox msk_todt 
            Height          =   255
            Left            =   3600
            TabIndex        =   4
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
         Begin MSMask.MaskEdBox msk_fromdt 
            Height          =   255
            Left            =   1200
            TabIndex        =   5
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
         Begin VB.Label lbl_fromdt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   2760
            TabIndex        =   7
            Top             =   120
            Width           =   705
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   0
            TabIndex        =   6
            Top             =   120
            Width           =   885
         End
      End
      Begin VB.OptionButton OPT_LAST 
         BackColor       =   &H80000000&
         Caption         =   "Last"
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
         Left            =   1320
         TabIndex        =   2
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton OPT_DATE 
         BackColor       =   &H80000000&
         Caption         =   "Date Wise"
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
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "cbcquery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SELMA106_FLAG As Boolean
Dim SELMODALITY_FLAG As Boolean
Dim SELMODALITY_QR As rdoQuery
Dim SELMODALITY_RS As rdoResultset
Dim STR As String
Dim DEPT_CODE As String

Dim selta101flag1 As Boolean
Dim selta101_qr1 As rdoQuery
Dim selta101_rs1 As rdoResultset
Dim SELDREQSTAFFDFLAG As Boolean
Dim SELDREQSTAFFD_QR As rdoQuery
Dim SELDREQSTAFFD_RS As rdoResultset

Public Sub grid_test()


End Sub


Private Sub cmb_dept_Click()
Dim CT1 As Integer
CT1 = 0
CT1 = InStr(Trim(cmb_dept.Text), "$")
CT1 = CT1 + 1
DEPT_CODE = Mid(Trim(cmb_dept.Text), CT1)

If SELMODALITY_FLAG = False Then
    SELMODALITY_QR(0) = Trim(DEPT_CODE)
    Set SELMODALITY_RS = SELMODALITY_QR.OpenResultset(1, 3)
    SELMODALITY_FLAG = True
Else
    SELMODALITY_QR(0) = Trim(DEPT_CODE)
    SELMODALITY_RS.MoreResults
    SELMODALITY_RS.Requery
End If

If SELMODALITY_RS.RowCount > 0 Then
    cmb_modality.clear
    While Not SELMODALITY_RS.EOF
        cmb_modality.AddItem Trim(SELMODALITY_RS("MODLITYDES")) & "$" & Trim(SELMODALITY_RS("MODALITYCD"))
        SELMODALITY_RS.MoveNext
    Wend
End If

End Sub

Private Sub Form_Load()
datetime

STR = "SELECT * FROM RISLIB.RMODALITY WHERE DEPT_CODE=?"
Set SELMODALITY_QR = rdb.CreateQuery("SEL_MODALITY", STR)

STR = "SELECT * FROM PATOOLSLIB.MA106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ' OR  DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='LZZ') "
Set SELMA106_QR = rdb.CreateQuery("SELMA1061_QRY", STR)

STR = "SELECT * FROM DISLIB.DREQSTAFFD WHERE EMPCODE=?"
Set SELDREQSTAFFD_QR = rdb.CreateQuery("SELSTAFF_QR", STR)

STR = "select * from patoolslib.ta101 where case_no=?"
Set selta101_qr1 = rdb.CreateQuery("selta101_qr", STR)


If SELMA106_FLAG = False Then
    Set SELMA106_RS = SELMA106_QR.OpenResultset(1, 3)
    SELMA106_FLAG = True
Else
    SELMA106_RS.MoreResults
    SELMA106_RS.Requery
End If

If SELMA106_RS.RowCount > 0 Then
cmb_dept.clear
    While Not SELMA106_RS.EOF
        cmb_dept.AddItem Trim(SELMA106_RS("DEPT_NAME")) & "$" & Trim(SELMA106_RS("DEPT_CODE"))
        SELMA106_RS.MoveNext
    Wend
End If

cmb_dept.Text = "HAEMATOPATHOLOGY LABORATORY$UZZ"
cmb_modality.Text = "HAEMATOLOGY$HL"
'msk_caseno.SetFocus

End Sub



Private Sub msk_caseno_LostFocus()
On Error GoTo MSG:

    If Len(Trim(msk_caseno.Text)) = 0 Then Exit Sub
    If Opt_patient.Value = True Then
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
            txt_name.Text = Trim(selta101_rs1("tag")) & " " & Trim(selta101_rs1("first_name")) & " " & Trim(selta101_rs1("mid_name")) & " " & Trim(selta101_rs1("last_name"))
            txt_sex.Text = Trim(selta101_rs1("sex"))
            txt_age.Text = DateDiff("YYYY", selta101_rs1("BIRTH_DATE"), CStr(DATE11))
        End If
        
    ElseIf opt_staff.Value = True Then
        STAFF
    End If
    
    Frame3.Enabled = True
    
    Exit Sub
    
MSG:
MsgBox Err.Description
End Sub

Private Sub Opt_patient_Click()
If Opt_patient.Value = True Then
    msk_caseno.Mask = "??>/#####"
    msk_caseno.Enabled = True
    msk_caseno.SetFocus
End If
End Sub

Private Sub opt_staff_Click()
If opt_staff.Value = True Then
    msk_caseno.Mask = "######"
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
        txt_name.Text = Trim(SELDREQSTAFFD_RS("PNAME"))
        txt_age.Text = Trim(SELDREQSTAFFD_RS("AGE"))
        txt_sex.Text = Trim(SELDREQSTAFFD_RS("SEX"))
    End If
End Sub
