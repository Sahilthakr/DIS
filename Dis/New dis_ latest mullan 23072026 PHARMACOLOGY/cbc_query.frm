VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form cbc_query 
   BackColor       =   &H80000018&
   Caption         =   "CBC QUERY"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10665
   LinkTopic       =   "Form2"
   ScaleHeight     =   12915
   ScaleWidth      =   23760
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   9120
      TabIndex        =   27
      Top             =   2760
      Width           =   1935
      Begin VB.CommandButton cmdback 
         BackColor       =   &H0088B7B7&
         Caption         =   "&BACK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdclear 
         BackColor       =   &H0088B7B7&
         Caption         =   "C&LEAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0088B7B7&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   600
      TabIndex        =   12
      Top             =   600
      Width           =   10575
      Begin VB.ComboBox CMB_NAME 
         Height          =   315
         Left            =   3120
         TabIndex        =   31
         Top             =   1440
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0088B7B7&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   840
         TabIndex        =   18
         Top             =   720
         Width           =   1335
         Begin VB.OptionButton Opt_patient 
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   20
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton opt_staff 
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   19
            Top             =   600
            Width           =   1095
         End
         Begin VB.Shape Shape1 
            BorderWidth     =   2
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   735
            Left            =   -120
            Shape           =   4  'Rounded Rectangle
            Top             =   120
            Width           =   1455
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   1440
         Width           =   855
      End
      Begin MSMask.MaskEdBox msk_caseno 
         Height          =   375
         Left            =   3240
         TabIndex        =   0
         Top             =   840
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   0
         PromptInclude   =   0   'False
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
      Begin VB.Shape Shape3 
         BorderColor     =   &H80000009&
         Height          =   1695
         Left            =   120
         Top             =   120
         Width           =   10335
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
      BackColor       =   &H0088B7B7&
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   600
      TabIndex        =   10
      Top             =   3960
      Width           =   10575
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4215
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   7435
         _Version        =   393216
         FixedRows       =   0
         BackColor       =   16777215
         BackColorFixed  =   8043708
         BackColorBkg    =   -2147483624
         GridColor       =   8421504
         GridLines       =   3
         AllowUserResizing=   3
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
      Begin VB.Shape Shape4 
         BackColor       =   &H80000006&
         BorderColor     =   &H80000009&
         Height          =   4455
         Left            =   120
         Top             =   120
         Width           =   10335
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   975
      Left            =   720
      TabIndex        =   1
      Top             =   2760
      Width           =   8295
      Begin VB.CommandButton CMD_EXPORT 
         BackColor       =   &H0088B7B7&
         Caption         =   "E&XPORT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmd_display 
         BackColor       =   &H0088B7B7&
         Caption         =   "D&ISPLAY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   0
         Width           =   1335
      End
      Begin VB.TextBox txt_num 
         Height          =   285
         Left            =   2160
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   4935
         Begin MSMask.MaskEdBox msk_todt 
            Height          =   255
            Left            =   3360
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
         Begin MSMask.MaskEdBox msk_fromdt 
            Height          =   255
            Left            =   960
            TabIndex        =   6
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
            Left            =   2520
            TabIndex        =   8
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
            TabIndex        =   7
            Top             =   120
            Width           =   885
         End
      End
      Begin VB.OptionButton OPT_LAST 
         BackColor       =   &H80000018&
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
         Left            =   600
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton OPT_DATE 
         BackColor       =   &H80000018&
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
         Left            =   600
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HOMI BHABHA CANCER HOSPITAL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00006633&
      Height          =   360
      Left            =   2715
      TabIndex        =   30
      Top             =   120
      Width           =   5445
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00008080&
      Height          =   1215
      Left            =   600
      Top             =   2640
      Width           =   10575
   End
End
Attribute VB_Name = "cbc_query"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SELMA106_FLAG As Boolean
Dim SELMODALITY_FLAG As Boolean
Dim SELMODALITY_QR As rdoQuery
Dim SELMODALITY_RS As rdoResultset
Dim str As String
Dim DEPT_CODE As String
Dim MOD_CODE As String

Dim var_date_f As String
Dim var_date_t As String
Dim req As String

Dim selta101flag1 As Boolean
Dim selta101_qr1 As rdoQuery
Dim selta101_rs1 As rdoResultset
Dim SELDREQSTAFFDFLAG As Boolean
Dim SELDREQSTAFFD_QR As rdoQuery
Dim SELDREQSTAFFD_RS As rdoResultset
Dim SELSTAFF_QR As rdoQuery
Dim SELSTAFF_RS As rdoResultset
Dim SELDHEMLFREQ_QR As rdoQuery
Dim SELDHEMLFREQ_RS As rdoResultset
Dim SELDHEMLF_QR As rdoQuery
Dim SELDHEMLF_RS As rdoResultset
Dim SELDHEMLFLAST_QR As rdoQuery
Dim SELDHEMLFLAST_RS As rdoResultset
Dim SELREQSTAFF_QR As rdoQuery
Dim SELREQSTAFF_RS As rdoResultset
Dim SELREQSTAFFLAST_QR As rdoQuery
Dim SELREQSTAFFLAST_RS As rdoResultset


Dim SELCBCTEST_QR As rdoQuery
Dim SELCBCTEST_RS As rdoResultset

Dim r As Integer
Dim c As Integer
Dim totalcnt As Integer

Public Sub GRID_DATA()
Dim GRIDCOL As Integer
Dim GRIDROW As Integer
r = 0
c = 0
GRIDCOL = MSFlexGrid1.Cols
GRIDROW = MSFlexGrid1.Rows

GRIDCOL = GRIDCOL - 1
GRIDROW = GRIDROW - 1

For c = 1 To GRIDCOL
    For r = 3 To GRIDROW
            
        SELDHEMLF_QR(0) = Trim(MSFlexGrid1.TextMatrix(0, c))
        SELDHEMLF_QR(1) = Trim(MSFlexGrid1.TextMatrix(r, 0))
        Set SELDHEMLF_RS = SELDHEMLF_QR.OpenResultset(1, 3)
        If SELDHEMLF_RS.EOF Then
            SELCBCTEST_QR(0) = Trim(MSFlexGrid1.TextMatrix(r, 0))
            SELCBCTEST_QR(1) = Trim(MSFlexGrid1.TextMatrix(r, 0))
            Set SELCBCTEST_RS = SELCBCTEST_QR.OpenResultset(1, 3)
            If SELCBCTEST_RS.EOF Then
            MSFlexGrid1.TextMatrix(r, c) = ""
            Else
             SELDHEMLF_QR(0) = Trim(MSFlexGrid1.TextMatrix(0, c))
             SELDHEMLF_QR(1) = Trim(SELCBCTEST_RS("FLD1"))
             Set SELDHEMLF_RS = SELDHEMLF_QR.OpenResultset(1, 3)
                If SELDHEMLF_RS.EOF Then
                MSFlexGrid1.TextMatrix(r, c) = ""
                Else
                MSFlexGrid1.TextMatrix(r, c) = Trim(SELDHEMLF_RS("OBSERVALUE"))
                End If
            End If
        Else
        MSFlexGrid1.TextMatrix(r, c) = Trim(SELDHEMLF_RS("OBSERVALUE"))
        End If
        
    Next r
Next c
End Sub
Public Sub GETDATA_LAST()

Dim num As Integer
num = Trim(txt_num.Text)
'NUM = NUM - 1
If Opt_patient.Value = True Then
    MSFlexGrid1.Cols = 1 + num

    SELDHEMLFLAST_QR(0) = Trim(msk_caseno.FormattedText)
    Set SELDHEMLFLAST_RS = SELDHEMLFLAST_QR.OpenResultset(1, 3)

    MSFlexGrid1.ColWidth(0) = 1800
    MSFlexGrid1.ColAlignment(0) = vbCenter
    For c = 1 To num
        MSFlexGrid1.ColWidth(c) = 1800
        MSFlexGrid1.ColAlignment(c) = vbCenter
        MSFlexGrid1.TextMatrix(0, c) = Trim(SELDHEMLFLAST_RS("REQNNO"))
        MSFlexGrid1.TextMatrix(1, c) = Trim(SELDHEMLFLAST_RS("FREPORTDT"))
        If Trim(SELDHEMLFLAST_RS("REMARKS")) = "" Then
        MSFlexGrid1.TextMatrix(2, c) = ""
        Else
        MSFlexGrid1.TextMatrix(2, c) = Trim(SELDHEMLFLAST_RS("REMARKS"))
        End If
               
        SELDHEMLFLAST_RS.MoveNext
    Next c
        
End If

If opt_staff.Value = True Then
    
    MSFlexGrid1.Cols = 1 + num
    
    SELREQSTAFFLAST_QR(0) = Trim(msk_caseno.FormattedText)
    SELREQSTAFFLAST_QR(1) = Trim(CMB_NAME.Text)
    
    Set SELREQSTAFFLAST_RS = SELREQSTAFFLAST_QR.OpenResultset(1, 3)
    
    MSFlexGrid1.ColWidth(0) = 1800
    MSFlexGrid1.ColAlignment(0) = vbCenter
    For c = 1 To num
        MSFlexGrid1.ColWidth(c) = 1800
        MSFlexGrid1.ColAlignment(c) = vbCenter
        MSFlexGrid1.TextMatrix(0, c) = Trim(SELREQSTAFFLAST_RS("REQNNO"))
        MSFlexGrid1.TextMatrix(1, c) = Trim(SELREQSTAFFLAST_RS("FREPORTDT"))
         If Trim(SELDHEMLFLAST_RS("REMARKS")) = "" Then
        MSFlexGrid1.TextMatrix(2, c) = ""
        Else
        MSFlexGrid1.TextMatrix(2, c) = Trim(SELDHEMLFLAST_RS("REMARKS"))
        End If
        SELREQSTAFFLAST_RS.MoveNext
    Next c
        
End If
End Sub


Public Sub GETDATA()
If Opt_patient.Value = True Then

    SELDHEMLFREQ_QR(0) = Trim(msk_caseno.FormattedText)
    SELDHEMLFREQ_QR(1) = msk_fromdt.FormattedText
    SELDHEMLFREQ_QR(2) = msk_todt.FormattedText

    Set SELDHEMLFREQ_RS = SELDHEMLFREQ_QR.OpenResultset(1, 3)

    MSFlexGrid1.Cols = 1 + SELDHEMLFREQ_RS.RowCount
    MSFlexGrid1.ColWidth(0) = 1800
    MSFlexGrid1.ColAlignment(0) = vbCenter
        
    If SELDHEMLFREQ_RS.RowCount > 0 Then

        For c = 1 To SELDHEMLFREQ_RS.RowCount
            MSFlexGrid1.ColWidth(c) = 1800
            MSFlexGrid1.ColAlignment(c) = vbCenter
            MSFlexGrid1.TextMatrix(0, c) = Trim(SELDHEMLFREQ_RS("REQNNO"))
            MSFlexGrid1.TextMatrix(1, c) = Trim(SELDHEMLFREQ_RS("FREPORTDT"))
            MSFlexGrid1.TextMatrix(2, c) = Trim(SELDHEMLFREQ_RS("REMARKS"))
            SELDHEMLFREQ_RS.MoveNext
        Next c
    Else
        MsgBox ("No CBC Report found for this patient."), , "Message"
        cmdclear_Click
        Opt_patient.Value = True
        msk_caseno.SetFocus
        Exit Sub
    End If
    
    GRID_DATA
    
End If
 
If opt_staff.Value = True Then
    SELREQSTAFF_QR(0) = Trim(msk_caseno.FormattedText)
    SELREQSTAFF_QR(1) = Trim(CMB_NAME.Text)
    SELREQSTAFF_QR(2) = msk_fromdt.FormattedText
    SELREQSTAFF_QR(3) = msk_todt.FormattedText
        
    Set SELREQSTAFF_RS = SELREQSTAFF_QR.OpenResultset(1, 3)
    
    MSFlexGrid1.Cols = 1 + SELREQSTAFF_RS.RowCount
    MSFlexGrid1.ColWidth(0) = 1800
    MSFlexGrid1.ColAlignment(0) = vbCenter
    
     If SELREQSTAFF_RS.RowCount > 0 Then
        For c = 1 To SELREQSTAFF_RS.RowCount
            MSFlexGrid1.ColWidth(c) = 1800
            MSFlexGrid1.ColAlignment(c) = vbCenter
            MSFlexGrid1.TextMatrix(0, c) = Trim(SELREQSTAFF_RS("REQNNO"))
            MSFlexGrid1.TextMatrix(1, c) = Trim(SELREQSTAFF_RS("FREPORTDT"))
            MSFlexGrid1.TextMatrix(2, c) = Trim(SELREQSTAFF_RS("REMARKS"))
            SELREQSTAFF_RS.MoveNext
        Next c
     Else
        MsgBox ("No CBC Report found for this patient."), , "Message"
        cmdclear_Click
        opt_staff.Value = True
        msk_caseno.SetFocus
        Exit Sub
     End If
    GRID_DATA
End If
End Sub


Public Sub GRID_DATEWISE()


If DEPT_CODE = "UZZ" Then
    If MOD_CODE = "HL" Then
            
            
            MSFlexGrid1.Rows = 32
            
            
            MSFlexGrid1.TextMatrix(0, 0) = "REQUISITN NO."
            MSFlexGrid1.TextMatrix(1, 0) = "REPORT DATE"
            MSFlexGrid1.TextMatrix(2, 0) = "MACHINE NAME"
            MSFlexGrid1.TextMatrix(3, 0) = "HAEMOGLOBIN"
            MSFlexGrid1.TextMatrix(4, 0) = "RBC COUNT"
            MSFlexGrid1.TextMatrix(5, 0) = "HAEMATOCRIT"
            MSFlexGrid1.TextMatrix(6, 0) = "MCV"
            MSFlexGrid1.TextMatrix(7, 0) = "MCH"
            MSFlexGrid1.TextMatrix(8, 0) = "MCHC"
            MSFlexGrid1.TextMatrix(9, 0) = "CHCM"
            MSFlexGrid1.TextMatrix(10, 0) = "HDW"
            MSFlexGrid1.TextMatrix(11, 0) = "CH"
            MSFlexGrid1.TextMatrix(12, 0) = "CHDW"
            MSFlexGrid1.TextMatrix(13, 0) = "PLATELETS"
            MSFlexGrid1.TextMatrix(14, 0) = "MPV"
            MSFlexGrid1.TextMatrix(15, 0) = "PDW"
            MSFlexGrid1.TextMatrix(16, 0) = "PCT"
            MSFlexGrid1.TextMatrix(17, 0) = "WBC COUNT"
            MSFlexGrid1.TextMatrix(18, 0) = "NEUTROPHIL %"
            MSFlexGrid1.TextMatrix(19, 0) = "LYMPHOCYTE %"
            MSFlexGrid1.TextMatrix(20, 0) = "MONOCYTE %"
            MSFlexGrid1.TextMatrix(21, 0) = "EOSINOPHIL %"
            MSFlexGrid1.TextMatrix(22, 0) = "BASOPHIL %"
            MSFlexGrid1.TextMatrix(23, 0) = "LARGE UNSTAINED CELLS % (LUC)"
            MSFlexGrid1.TextMatrix(24, 0) = "NRBCS % (NUCLEATED RBCS)"
            MSFlexGrid1.TextMatrix(25, 0) = "NEUTROPHIL ABS"
            MSFlexGrid1.TextMatrix(26, 0) = "LYMPHOCYTE ABS"
            MSFlexGrid1.TextMatrix(27, 0) = "MONOCYTE ABS"
            MSFlexGrid1.TextMatrix(28, 0) = "EOSINOPHIL ABS"
            MSFlexGrid1.TextMatrix(29, 0) = "BASOPHIL ABS"
            MSFlexGrid1.TextMatrix(30, 0) = "LARGE UNSTAINED CELLS (LUC) ABS"
            MSFlexGrid1.TextMatrix(31, 0) = "NUCLEATED RBCS(NRBCS) ABS"
            MSFlexGrid1.FixedRows = 3
           
'             Set SELHLTEST_RS = SELHLTEST_QR.OpenResultset(1, 3)
'            MSFlexGrid1.Rows = 2 + SELHLTEST_RS.RowCount
'            MSFlexGrid1.FixedRows = 2
'            MSFlexGrid1.FixedCols = 1
'
'            MSFlexGrid1.TextMatrix(0, 0) = "REQUISITN NO."
'            MSFlexGrid1.TextMatrix(1, 0) = "REPORT DATE"
'
'            r = 2
'
'            While Not SELHLTEST_RS.EOF
'               MSFlexGrid1.TextMatrix(r, 0) = Trim(SELHLTEST_RS(0))
'               r = r + 1
'               SELHLTEST_RS.MoveNext
'            Wend
                        
    End If
End If
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

Private Sub cmb_modality_LostFocus()
Dim CT2 As Integer
CT2 = 0
CT2 = InStr(Trim(cmb_modality.Text), "$")
CT2 = CT2 + 1
MOD_CODE = Mid(Trim(cmb_modality.Text), CT2)
End Sub

Private Sub CMB_NAME_Click()

SELSTAFF_QR(0) = Trim(msk_caseno.FormattedText)
SELSTAFF_QR(1) = Trim(CMB_NAME.Text)

Set SELSTAFF_RS = SELSTAFF_QR.OpenResultset(1, 3)
If SELSTAFF_RS.RowCount > 0 Then
    'While Not SELSTAFF_RS.EOF
        txt_Age.Text = Trim(SELSTAFF_RS("AGE"))
        txt_sex.Text = Trim(SELSTAFF_RS("SEX"))
    'Wend
End If

Frame3.Enabled = True
End Sub

Private Sub cmd_display_Click()

Dim CNT As Integer
If OPT_LAST.Value = True Then

        If txt_num.Text = "" Then
        Else
            CNT = Trim(txt_num.Text)
        End If

    If txt_num.Text = "" Then
        MsgBox ("Please enter number..."), , "Message"
        txt_num.SetFocus
        Exit Sub
    Else
      
          If CNT = 0 Then
             txt_num.Text = ""
             txt_num.SetFocus
             Exit Sub
         ElseIf CNT > totalcnt Then
           MsgBox ("Report number can not be greater than total report"), , "Message"
            txt_num.Text = ""
            txt_num.SetFocus
            Exit Sub
           End If
    End If
End If
If OPT_DATE.Value = True Then
    var_date_f = Year(msk_fromdt.FormattedText) & "," & month(msk_fromdt.FormattedText) & "," & Day(msk_fromdt.FormattedText)
    ar_date_t = Year(msk_todt.FormattedText) & "," & month(msk_todt.FormattedText) & "," & Day(msk_todt.FormattedText)

    GETDATA
    
ElseIf OPT_LAST.Value = True Then
    GETDATA_LAST
    GRID_DATA
End If
End Sub

Private Sub CMD_EXPORT_Click()
Dim i As Integer
Dim GRID_ROW As Integer
Dim CH As String
Dim Cl As String
Dim INC As Integer

CH = MsgBox("Do u wish to transfer data to an Excel sheet ...?", vbYesNo, "Message")

If CH = vbYes Then
    Dim xlOBJ As Excel.Application
    Dim xlWB As Excel.Workbook
    Dim xlSH As Excel.Worksheet

    Set xlOBJ = New Excel.Application

    Set xlWB = xlOBJ.Workbooks.add

    GRID_ROW = MSFlexGrid1.Rows
    GRID_ROW = GRID_ROW - 1

    Set xlSH = xlWB.Worksheets(1)
    xlSH.Activate
    INC = 8

        With xlOBJ.ActiveWorkbook.ActiveSheet
            .range("B2") = "CASE NO : " + msk_caseno
            If Opt_patient.Value = True Then
                  .range("B4") = "PATIENT NAME : " + txt_name
            ElseIf opt_staff.Value = True Then
                  .range("B4") = "PATIENT NAME : " + CMB_NAME
            End If

            .range("B5") = "AGE : " + txt_Age
            .range("D5") = "SEX : " + txt_sex

            If OPT_DATE.Value = True Then
                .range("B7") = "From Date : " + msk_fromdt.FormattedText + "  To Date : " + msk_todt.FormattedText
            End If

        End With

    For i = 0 To GRID_ROW
        Clipboard.clear

        With MSFlexGrid1
            .col = 0
            .row = i
            .ColSel = .Cols - 1
            .RowSel = i
            Clipboard.SetText .Clip
        End With
        INC = INC + 1
        Cl = "A" & INC

        With xlOBJ.ActiveWorkbook.ActiveSheet


            .range(Cl).Select
            .Paste

        End With
    Next i

    xlOBJ.Visible = True

End If


End Sub

Private Sub cmdback_Click()
Me.MousePointer = 11
Me.Hide
DIS_MAIN.show
Me.MousePointer = 0
End Sub

Private Sub cmdclear_Click()
Opt_patient.Value = False
opt_staff.Value = False
OPT_DATE.Value = False
OPT_LAST.Value = False
msk_caseno.Text = ""
txt_name.Text = ""
CMB_NAME.Text = ""
txt_Age.Text = ""
txt_sex.Text = ""
Frame3.Enabled = False
Frame6.Visible = False
txt_num.Visible = False
DEPT_CODE = ""
MOD_CODE = ""

MSFlexGrid1.clear

MSFlexGrid1.Rows = 4
MSFlexGrid1.Cols = 2
MSFlexGrid1.FixedRows = 3
MSFlexGrid1.FixedCols = 1
End Sub

Private Sub Form_Load()
datetime

str = "SELECT * FROM RISLIB.RMODALITY WHERE DEPT_CODE=?"
Set SELMODALITY_QR = rdb.CreateQuery("SEL_MODALITY", str)

str = "SELECT * FROM PATOOLSLIB.MA106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ' OR  DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='LZZ') "
Set SELMA106_QR = rdb.CreateQuery("SELMA1061_QRY", str)


str = "SELECT DISTINCT(PNAME) FROM DISLIB.DREQSTAFFD WHERE EMPCODE=?"
Set SELDREQSTAFFD_QR = rdb.CreateQuery("SELSTAFF_QR", str)

str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE EMPCODE=? AND PNAME=? ORDER BY AGE DESC"
Set SELSTAFF_QR = rdb.CreateQuery("SELSTAFF", str)

str = "select * from patoolslib.ta101 where case_no=?"
Set selta101_qr1 = rdb.CreateQuery("selta101_qr", str)
'SHWETA 25-07-2012
str = "select TESTNAME,FLD1 from DISLIB.CBCTEST where TESTNAMe=? OR  FLD1=?"
Set SELCBCTEST_QR = rdb.CreateQuery("", str)

'SHWETA 25-07-2012
str = "SELECT DISTINCT REQNNO,FREPORTDT,remarks FROM DISLIB.DHEMLF WHERE CASENO=? AND FREPORTDT >= ? AND FREPORTDT <=? and (reqnstat='T' or reqnstat='F' ) ORDER BY FREPORTDT DESC"
Set SELDHEMLFREQ_QR = rdb.CreateQuery("SELDHEMLFREQ", str)

str = "SELECT OBSERVALUE FROM DISLIB.DHEMLF WHERE REQNNO=? AND PRO_CODE=?"
Set SELDHEMLF_QR = rdb.CreateQuery("SELDHEMLF", str)
'SHWETA 25-07-2012
str = "SELECT DISTINCT REQNNO,FREPORTDT,remarks FROM DISLIB.DHEMLF WHERE CASENO=? and (reqnstat='T' or reqnstat='F' ) ORDER BY FREPORTDT DESC"
Set SELDHEMLFLAST_QR = rdb.CreateQuery("SELDHEMLFLAST", str)

str = "SELECT DISTINCT REQNNO,FREPORTDT FROM DISLIB.DREQSTAFFD T1,DISLIB.DHEMLF T2 WHERE T1.EMPCODE=? AND T1.PNAME = ? AND T2.FREPORTDT>=? AND T2.FREPORTDT<=? AND T1.REQNO LIKE 'UZZHL%' AND T1.REQNO=T2.REQNNO AND T2.REQNSTAT='F' ORDER BY T2.REQNNO DESC"
Set SELREQSTAFF_QR = rdb.CreateQuery("SELREQSTAFF", str)

str = "SELECT DISTINCT REQNNO,FREPORTDT FROM DISLIB.DREQSTAFFD T1,DISLIB.DHEMLF T2 WHERE T1.EMPCODE=? AND T1.PNAME = ? AND T1.REQNO LIKE 'UZZHL%' AND T1.REQNO=T2.REQNNO AND T2.REQNSTAT='F' ORDER BY T2.REQNNO DESC"
Set SELREQSTAFFLAST_QR = rdb.CreateQuery("SELREQSTAFFLAST", str)

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

'OPT_DATE.Enabled = False
'OPT_LAST.Enabled = False
'Opt_patient.Value = True
DEPT_CODE = ""
MOD_CODE = ""
req = ""

End Sub






Private Sub msk_caseno_GotFocus()
Dim CT3 As Integer
CT3 = 0
CT3 = InStr(Trim(cmb_modality.Text), "$")
CT3 = CT3 + 1
MOD_CODE = Mid(Trim(cmb_modality.Text), CT3)

CT3 = 0
CT3 = InStr(Trim(cmb_dept.Text), "$")
CT3 = CT3 + 1
DEPT_CODE = Mid(Trim(cmb_dept.Text), CT3)

req = ""
req = DEPT_CODE & MOD_CODE

msk_caseno.SetFocus
End Sub

Private Sub msk_caseno_LostFocus()
On Error GoTo MSG:

    If Len(Trim(msk_caseno.Text)) = 0 Then Exit Sub
    If Opt_patient.Value = True Then
'        CASENO = padzero(5, "C", msk_caseno.Text, 3)
'        msk_caseno.Text = UCase(CASENO)   ''commented by Monika on 05/12/2022
'''' Commented PADZERO FUNCTION INSTEAD USED NEWNOPAD FUNCTION
'cno_pad = padzero(5, "C", msk_caseno.Text, 3)
    

    
  If IsNumeric(Left(Trim(msk_caseno.Text), 1)) Then
          Dim Ks As Integer
        Dim str1s As String
        Dim strs As String
            strs = Trim(msk_caseno.Text)
            Ks = Len(strs)
            str1s = Mid(strs, 5, Ks)
              Dim sel_CASE_rs  As rdo.rdoResultset
              Set sel_CASE_rs = rdb.OpenResultset(" SELECT * FROM PATOOLSLIB.NEWCASETBL WHERE CASE_CODE = '" & Left(UCase(msk_caseno.FormattedText), 4) & "' ", 1, 1)
              If Not sel_CASE_rs.EOF Then
                    Dim CSNO As String
                    CSNO = Trim(sel_CASE_rs(1))
                    CSNO = CSNO & str1s
                     cno_pad = NEWNOpad(CSNO)
              Else
               cno_pad = NEWNOpad(UCase(msk_caseno))
              End If
    
  Else
    cno_pad = NEWNOpad(UCase(msk_caseno))
  End If
     msk_caseno.Text = cno_pad
''''''''''''''''''''till here (01/12/2022)'''''''''


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
            txt_Age.Text = DateDiff("YYYY", selta101_rs1("BIRTH_DATE"), CStr(DATE11))
        End If
        
        Frame3.Enabled = True
        
    ElseIf opt_staff.Value = True Then
        STAFF
    End If
    
    
    Frame5.Enabled = True
    
    
    Exit Sub
    
MSG:
MsgBox Err.Description
End Sub

Private Sub msk_fromdt_LostFocus()
    msk_todt.SetFocus
End Sub



Private Sub OPT_DATE_Click()
If OPT_DATE.Value = True Then
    Frame6.Visible = True
    txt_num.Visible = False
    OPT_LAST.Value = False
    GRID_DATEWISE
    msk_fromdt.Text = DATE11
    msk_todt.Text = DATE11
    msk_fromdt.SetFocus
    
End If
End Sub

Private Sub OPT_LAST_Click()
'Dim TOTALCNT As Integer
totalcnt = 0
If OPT_LAST.Value = True Then
    Frame6.Visible = False
    txt_num.Text = ""
    txt_num.Visible = True
    txt_num.Enabled = True
    GRID_DATEWISE
    txt_num.SetFocus
    
    If Opt_patient.Value = True Then
            SELDHEMLFLAST_QR(0) = Trim(msk_caseno.FormattedText)
            Set SELDHEMLFLAST_RS = SELDHEMLFLAST_QR.OpenResultset(1, 3)
            totalcnt = SELDHEMLFLAST_RS.RowCount
            If totalcnt = 0 Then
                MsgBox ("No requisitions for this patient."), , "Message"
                cmdclear_Click
                Opt_patient.Value = True
                msk_caseno.SetFocus
            Else
                MsgBox ("The number of CBC requisitions for this patient :  ") & SELDHEMLFLAST_RS.RowCount, , "Message"
            End If
    ElseIf opt_staff.Value = True Then
            
            SELREQSTAFFLAST_QR(0) = Trim(msk_caseno.FormattedText)
            SELREQSTAFFLAST_QR(1) = Trim(CMB_NAME.Text)
            Set SELREQSTAFFLAST_RS = SELREQSTAFFLAST_QR.OpenResultset(1, 3)
            totalcnt = SELREQSTAFFLAST_RS.RowCount
            If totalcnt = 0 Then
                MsgBox ("No requisitions for this patient."), , "Message"
                cmdclear_Click
                opt_staff.Value = True
                msk_caseno.SetFocus
            Else
                MsgBox ("The number of requisitions for this patient :  ") & SELREQSTAFFLAST_RS.RowCount, , "Message"
            End If
    End If
End If
    

'    SELDHEMCBCDET_QR1(0) = Trim(msk_caseno.FormattedText)
'    Set SELDHEMCBCDET_RS1 = SELDHEMCBCDET_QR1.OpenResultset(1, 3)
'
'    If SELDHEMCBCDET_RS1.RowCount > 0 Then
'        While Not SELDHEMCBCDET_RS1.EOF
'                SELREQINDHEM_QR(0) = SELDHEMCBCDET_RS1("REQNNO")
'                Set SELREQINDHEM_RS = SELREQINDHEM_QR.OpenResultset(1, 3)
'
'
'            If SELREQINDHEM_RS.RowCount > 0 Then
'                TOTALCNT = TOTALCNT + 1
'            End If
'            SELDHEMCBCDET_RS1.MoveNext
'        Wend
'    End If
'    'txt_num.Text = TOTALCNT
'
'        MsgBox "Total number of REQUISITIONS : " & TOTALCNT, , "DIS"
'    txt_num.SetFocus
'End If
End Sub

Private Sub Opt_patient_Click()
If Opt_patient.Value = True Then
   '' msk_caseno.Mask = "??>/#####" ''commented by Monika on 05/12/2022
    msk_caseno.Enabled = True
    msk_caseno.SetFocus
    txt_name.Visible = True
    CMB_NAME.Visible = False
    txt_name.Text = ""
    txt_Age.Text = ""
    txt_sex.Text = ""
End If
End Sub

Private Sub opt_staff_Click()
If opt_staff.Value = True Then
    msk_caseno.Mask = "######"
    msk_caseno.Enabled = True
    msk_caseno.SetFocus
    CMB_NAME.Visible = True
    txt_name.Visible = False
    CMB_NAME.clear
    txt_Age.Text = ""
    txt_sex.Text = ""
End If
End Sub

Public Sub STAFF()
    
  CMB_NAME.clear
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
        While Not SELDREQSTAFFD_RS.EOF
            CMB_NAME.AddItem (SELDREQSTAFFD_RS("PNAME"))
            SELDREQSTAFFD_RS.MoveNext
        Wend
    End If
End Sub


Private Sub txt_num_KeyPress(KeyAscii As Integer)
Dim inx As Integer
inx = KeyAscii

If inx <> 8 And inx <> 47 And inx <> 48 And inx <> 49 And inx <> 50 And inx <> 51 And inx <> 52 And inx <> 53 And inx <> 54 And inx <> 55 And inx <> 56 And inx <> 57 Then
    MsgBox "Please Enter Numeric value", , "Message"
    txt_num.Text = ""
    txt_num.SetFocus
End If
End Sub


