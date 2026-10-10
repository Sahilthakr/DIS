VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_PathLblPrint 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Pathology Barcode Label Print"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13200
   LinkTopic       =   "Form2"
   ScaleHeight     =   8970
   ScaleWidth      =   13200
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chk_block 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Block"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6240
      TabIndex        =   35
      Top             =   5280
      Width           =   975
   End
   Begin VB.CheckBox chk_slide 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Slide"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5280
      TabIndex        =   34
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox TXTPATHG 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   0
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txt_printQty 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   4800
      Width           =   1095
   End
   Begin VB.TextBox txt_reqno 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      TabIndex        =   31
      Top             =   3240
      Width           =   2295
   End
   Begin VB.TextBox txt_seqno 
      Height          =   285
      Left            =   3480
      TabIndex        =   4
      Top             =   6240
      Width           =   1575
   End
   Begin VB.TextBox txt_pathno 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Height          =   975
      Left            =   960
      TabIndex        =   21
      Top             =   7920
      Width           =   7935
      Begin VB.CommandButton cmd_exit 
         BackColor       =   &H0080C0FF&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         MaskColor       =   &H0080C0FF&
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmb_view 
         BackColor       =   &H0080C0FF&
         Caption         =   "&View"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         MaskColor       =   &H0080C0FF&
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmd_cancel 
         BackColor       =   &H0080C0FF&
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         MaskColor       =   &H0080C0FF&
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmb_print 
         BackColor       =   &H0080C0FF&
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         MaskColor       =   &H0080C0FF&
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmd_back 
         BackColor       =   &H0080C0FF&
         Caption         =   "&Back"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         MaskColor       =   &H0080C0FF&
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.TextBox TXT_RCVDTTM 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   19
      Top             =   3000
      Width           =   3495
   End
   Begin VB.TextBox TXT_COLLDTTM 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   18
      Top             =   2520
      Width           =   3495
   End
   Begin VB.TextBox txt_name 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1320
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3960
      Width           =   5535
   End
   Begin VB.TextBox txt_sex 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1800
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4440
      Width           =   375
   End
   Begin VB.TextBox txt_Age 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3000
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox txt_catgory 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5280
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1455
   End
   Begin VB.ComboBox CMB_FZZLOCATIONS 
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "frm_PathLblPrint.frx":0000
      Left            =   5040
      List            =   "frm_PathLblPrint.frx":0002
      TabIndex        =   2
      Top             =   1680
      Width           =   2655
   End
   Begin MSMask.MaskEdBox msk_caseno 
      Height          =   375
      Left            =   1320
      TabIndex        =   26
      ToolTipText     =   "ENTER THE CASE NUMBER TO DISPLAY PATIENT DETAILS"
      Top             =   2520
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      ForeColor       =   0
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      HideSelection   =   0   'False
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2055
      Left            =   2760
      TabIndex        =   28
      Top             =   5880
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   3625
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   8438015
      ForeColorFixed  =   128
      GridColor       =   0
      AllowUserResizing=   3
      FormatString    =   "S.No |^ Sequence No."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Print for"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   33
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Print Quantity"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   32
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "ReqNo."
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   30
      Top             =   3240
      Width           =   855
   End
   Begin VB.Shape Shape1 
      Height          =   3375
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   2400
      Width           =   9615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "HOMI BHABHA CANCER HOSPITAL AND RESEARCH CENTRE"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   29
      Top             =   240
      Width           =   9735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Sample Received Date/Time"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   20
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Sample Collection Date/Time"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   17
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lbl_age 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2640
      TabIndex        =   16
      Top             =   4440
      Width           =   360
   End
   Begin VB.Label lbl_name 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   600
      TabIndex        =   15
      Top             =   3960
      Width           =   525
   End
   Begin VB.Label lbl_sex 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1320
      TabIndex        =   14
      Top             =   4440
      Width           =   330
   End
   Begin VB.Label lbl_catgory 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   4320
      TabIndex        =   13
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Case No."
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Sample Collection Unit"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Path No"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PATHOLOGY BARCODE LABEL PRINT"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   720
      Width           =   3255
   End
End
Attribute VB_Name = "frm_PathLblPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim selsclqr As rdoQuery
Dim selsclrs As rdoResultset

Dim seldrequisitn_qr As rdoQuery
Dim seldrequisitn_rs As rdoResultset

Dim seldrequisitn_qrC As rdoQuery
Dim seldrequisitn_rsC As rdoResultset

Dim SELADDENDUMCYQR As rdoQuery
Dim SELADDENDUMCYRS As rdoResultset

Dim SELRREQNO_QR As rdoQuery
Dim SELRREQNO_RS As rdoResultset

Dim sel_rtime_qr As rdoQuery
Dim sel_rtime_rs As rdoResultset

Dim selta101qr As rdoQuery
Dim selta101rs As rdoResultset

Dim selta1012_qr As rdoQuery
Dim selta1012_rs As rdoResultset

Dim SELMA301_QR As rdoQuery
Dim SELMA301_RS As rdoResultset

Dim SELDREQNOQR As rdoQuery
Dim SELDREQNORS As rdoResultset

Dim INS_PATHBRCLBL_QR As rdoQuery

Dim str As String
Dim prevseqno As Integer
Function CHECKDUPLICATE() As Boolean
Dim SELPATHBRCLBL_QR As rdoQuery
Dim SELPATHBRCLBL_RS As rdoResultset
Dim STR11 As String
STR11 = "SELECT * FROM DISLIB.PATHBRCLBL WHERE PATHNO=?"
Set SELPATHBRCLBL_QR = rdb.CreateQuery("", STR11)
SELPATHBRCLBL_QR(0) = Trim(txt_pathno.Text)

Set SELPATHBRCLBL_RS = SELPATHBRCLBL_QR.OpenResultset(1, 3)
If SELPATHBRCLBL_RS.RowCount = 0 Then
CHECKDUPLICATE = False
Else
CHECKDUPLICATE = True
End If
End Function
Sub SAVEDATA()
If CHECKDUPLICATE = True Then
Exit Sub
End If


Dim MSG, MSG1 As String
MSG = ""
MSG1 = ""
For i = 1 To MSFlexGrid1.Rows - 1
        MSG = UCase(Trim(MSFlexGrid1.TextMatrix(i, 1)))
        If Trim(MSG1) = "" Then
        MSG1 = MSG
        Else
        MSG1 = MSG1 & "," & MSG
        End If
Next
        INS_PATHBRCLBL_QR(0) = Trim(txt_reqno.Text)
        INS_PATHBRCLBL_QR(1) = Trim(msk_caseno.FormattedText)
        INS_PATHBRCLBL_QR(2) = txt_pathno.Text
        INS_PATHBRCLBL_QR(3) = CMB_FZZLOCATIONS.Text
        INS_PATHBRCLBL_QR(4) = Trim(txt_catgory.Text)
        INS_PATHBRCLBL_QR(5) = txt_printQty.Text
        INS_PATHBRCLBL_QR(6) = MSG1
        INS_PATHBRCLBL_QR.Execute
End Sub
Private Sub cmb_print_Click()




If Trim(msk_caseno.Text) = "" Then
MsgBox "Case no is Blank", vbInformation
Exit Sub
End If

If Trim(txt_pathno.Text) = "" Then
MsgBox "Path no is Blank", vbInformation
Exit Sub
End If


Dim printreason As String
Dim lprng As Integer
If chk_slide.Value = 1 And chk_block.Value = 1 Then
printreason = "12"
lprng = 2
Else
    If chk_slide.Value = 1 Then
    printreason = "1"
    lprng = 1
    End If
    If chk_block = 1 Then
    printreason = "2"
    lprng = 1
    End If
End If
Dim nlp As Integer
For nlp = 1 To lprng






Next
'
Dim FSTEXT As String
Dim pq, LP, l, r, ln1 As Integer
Dim D1, rm1, bctxt, bctxt1 As String
pq = txt_printQty.Text
If pq > 0 Then
'SAVEDATA
f1 = FreeFile()
Open "c:\DISREPORTSEC\DATA1\micro.TXT" For Output As #f1
For LP = 1 To pq
PRNTSTR1 = ""
'PRNTSTR1 = "TEC BARCODE PRINTER JOB START" & vbCrLf
PRNTSTR1 = PRNTSTR1 & "{D0221,0225,0201|}" & vbCrLf '{D0221,0225,0201|}
PRNTSTR1 = PRNTSTR1 & "{C|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{XB00;0019,0050,A,3,02,0,0047,+0000000000,000,0,00|}" '& Trim(msk_caseno.FormattedText) & "|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{PC000;0080,0040,05,05,J,00,B|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{LC;0080,0041,0167,0041,0,1|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{PC001;0032,0174,1,1,C,00,B|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{PC002;0102,0174,05,1,R,00,B|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{PC003;0047,0019,05,05,J,00,B|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{LC;0047,0021,0199,0021,0,1|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{PC004;0018,0136,1,1,D,00,B|}" & vbCrLf



bctxt1 = ""
 bctxt = LP
' ln1 = bctxt
 Dim CHKNO As Integer
''''''''''''
 For i = 1 To MSFlexGrid1.Rows - 1
 
 
        If IsNumeric(Left(Trim(MSFlexGrid1.TextMatrix(i, 1)), 2)) = True Then
            ln1 = Left(Trim(MSFlexGrid1.TextMatrix(i, 1)), 2)
        Else
            If IsNumeric(Left(Trim(MSFlexGrid1.TextMatrix(i, 1)), 1)) = True Then
            ln1 = Left(Trim(MSFlexGrid1.TextMatrix(i, 1)), 1)
            End If
        End If
        
If ln1 = LP Then
D1 = Trim(MSFlexGrid1.TextMatrix(i, 1))

 If Trim(D1) <> "" Then
  
    FSTEXT = Trim(D1)
    If Trim(FSTEXT) <> "" Then
      Arr = Split(FSTEXT, " ")
      bctxt = Arr(0)
      
        If UBound(Arr) >= 2 Then
        bctxt1 = Arr(1) & " " & Arr(2)
        ElseIf UBound(Arr) = 1 Then
        bctxt1 = Arr(1)
        Else
        bctxt1 = ""
        End If
        
      
    End If
   Exit For
Else
    bctxt = LP
   bctxt1 = ""
End If
 

End If




 Next
PRNTSTR1 = PRNTSTR1 & "{XB00;0025,0050,A,3,02,0,0047,+0000000000,000,0,00=>5" & Left(Trim(txt_pathno.Text), 6) & ">6" & Right(Trim(txt_pathno.Text), 2) & "|}" & vbCrLf
PRNTSTR1 = PRNTSTR1 & "{PC000;0104,0042,05,05,E,00,B=" & Trim(msk_caseno.FormattedText) & "|}" & vbCrLf
PRNTSTR1 = PRNTSTR1 & "{PC001;0027,0172,05,05,E,00,B=" & Trim(bctxt) & "|}" & vbCrLf
PRNTSTR1 = PRNTSTR1 & "{PC002;0089,0175,05,1,R,00,B=" & Trim(bctxt1) & "|}" & vbCrLf
PRNTSTR1 = PRNTSTR1 & "{PC003;0066,0019,05,05,J,00,B=TMH|}" & vbCrLf
PRNTSTR1 = PRNTSTR1 & "{LC;0066,0021,0114,0021,0,1|}" & vbCrLf
If Trim(TXTPATHG.Text) = "G" Then
'PRNTSTR1 = PRNTSTR1 & "{RB00;>6G>5" & Replace(Left(Trim(txt_pathno.Text), Len(txt_pathno.Text) - 2), "/", "") & ">6" & Trim(Right(Trim(txt_pathno.Text), 2)) & "|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{RC004G;" & Trim(txt_pathno.Text) & "|}" & vbCrLf
PRNTSTR1 = PRNTSTR1 & "{PC004;0005,0136,1,1,D,00,B=G" & Replace(Trim(txt_pathno.Text), "/", "") & "|}" & vbCrLf
Else
'PRNTSTR1 = PRNTSTR1 & "{RB00;>6G>5" & Replace(Left(Trim(txt_pathno.Text), Len(txt_pathno.Text) - 2), "/", "") & ">6" & Trim(Right(Trim(txt_pathno.Text), 2)) & "|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{RC004;" & Trim(txt_pathno.Text) & "|}" & vbCrLf
PRNTSTR1 = PRNTSTR1 & "{PC004;0005,0136,1,1,D,00,B=" & Replace(Trim(txt_pathno.Text), "/", "") & "|}" & vbCrLf
End If
PRNTSTR1 = PRNTSTR1 & "{XS;I,0001,0002C4200|} " & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{RC004;" & Trim(txt_pathno.Text) & "|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{RC000;Hospital|}" & vbCrLf
'p 'RNTSTR1 = PRNTSTR1 & "{RC001;" & Trim(bctxt) & "|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{RC002;" & Trim(bctxt1) & "|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{RC003;Tata Memorial|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{XS;I,0001,0002C4201|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "TEC BARCODE PRINTER PAGE END" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "TEC BARCODE PRINTER JOB END" & vbCrLf

Print #f1, PRNTSTR1
Next
Close #f1
Shell "c:\DISREPORTSEC\DATA1\micro.bat", 0

SAVEDATA

End If
  

End Sub

Private Sub Form_Load()

MSFlexGrid1.Rows = 2
prevseqno = 0
str = "SELECT * FROM DISLIB.DREQUISITN WHERE PATHNO=?"
Set seldrequisitn_qr = rdb.CreateQuery("", str)

str = "SELECT * FROM DISLIB.DREQUISITN WHERE CASENO=?"
Set seldrequisitn_qrC = rdb.CreateQuery("", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
Set SELRREQNO_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM RISLIB.RTIMELOG WHERE REQNNO=?"
Set sel_rtime_qr = rdb.CreateQuery("", str)

str = "SELECT * FROM PATOOLSLIB.TA101 WHERE CASE_NO=?"
Set selta101qr = rdb.CreateQuery("", str)

str = "select * from patoolslib.ta1012 where case_no=?"
Set selta1012_qr = rdb.CreateQuery("", str)

str = "SELECT * FROM PATOOLSLIB.MA301 WHERE SITECD=?"
Set SELMA301_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO=?"
Set SELDREQNOQR = rdb.CreateQuery("", str)

str = "INSERT INTO DISLIB.PATHBRCLBL(REQNO,CASENO,PATHNO,COLLUNIT,PAT_CATG,PRNTQTY,SEQNO) VALUES(?,?,?,?,?,?,?)"
Set INS_PATHBRCLBL_QR = rdb.CreateQuery("", str)


str = "SELECT LOCATION FROM DISLIB.SAMCOLLOCM "
Set selsclqr = rdb.CreateQuery("", str)
Set selsclrs = selsclqr.OpenResultset(1, 3)
If selsclrs.RowCount > 0 Then
selsclrs.MoveFirst
Dim KL As Integer
    For KL = 0 To selsclrs.RowCount - 1
        CMB_FZZLOCATIONS.AddItem (Trim(selsclrs("LOCATION")))
        selsclrs.MoveNext
    Next
Else
MsgBox "DATA NOT FOUND IN SAMPLE COLLECTION LOCATION MASTER", vbExclamation
End If


txt_pathno.Text = ""
txt_name.Text = ""

txt_pathno.Enabled = False
End Sub


Private Sub cmb_view_Click()
cmd_cancel.Enabled = True
cmb_print.Enabled = True
cmd_back.Enabled = True
cmb_view.Enabled = False

txt_pathno.Enabled = True
txt_pathno.Text = ""
txt_name.Text = ""
TXTPATHG.Text = ""
'txt_pathno.SetFocus
TXTPATHG.SetFocus
End Sub

Private Sub cmd_cancel_Click()
prevseqno = 0
cmd_cancel.Enabled = False
cmb_print.Enabled = False
cmd_back.Enabled = True
cmb_view.Enabled = True

txt_pathno.Enabled = False

MSFlexGrid1.row = 1
MSFlexGrid1.col = 1
txt_seqno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top

txt_seqno.Visible = False
MSFlexGrid1.clear

txt_pathno.Text = ""
txt_name.Text = ""
msk_caseno.Text = ""
txt_Age.Text = ""
txt_sex.Text = ""
txt_catgory.Text = ""
TXT_COLLDTTM.Text = ""
TXT_RCVDTTM.Text = ""
txt_printQty.Text = ""
txt_reqno.Text = ""
TXTPATHG.Text = ""
prevseqno = 0
cmb_view.Enabled = True
'cmb_view_Click
End Sub

Private Sub cmd_back_Click()
cmd_cancel.Enabled = False
cmb_print.Enabled = False
cmd_back.Enabled = True
cmb_view.Enabled = True
txt_pathno.Text = ""
txt_name.Text = ""
'msf_reqstat.Rows = 1
Me.Hide
DIS_MAIN.show
End Sub

Private Sub cmd_exit_Click()
End
End Sub






Private Sub txt_pathno_LostFocus()
txt_printQty.Text = ""
txt_reqno.Text = ""
If Trim(txt_pathno.Text) = "" Then
    MsgBox "Path No. Should Not Be Empty"
    Exit Sub
End If

''shweta for / in pathno
Dim sl As Integer
sl = InStr(txt_pathno.Text, "/")
Dim P1, p2, p3, p4 As String
If sl = 0 Then
P1 = Trim(txt_pathno.Text)
p2 = Right(Trim(P1), 2)
p3 = Left(Trim(P1), Len(P1) - 2)
p4 = p3 & "/" & UCase(p2)
Else
p4 = Trim(txt_pathno.Text)
End If

txt_pathno.Enabled = False
txt_seqno.Visible = True
MSFlexGrid1.row = 1
MSFlexGrid1.col = 1
txt_seqno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top

MSFlexGrid1.Rows = 2
MSFlexGrid1.FormatString = " Sr.No |^Sequence No."

seldrequisitn_qr(0) = UCase(Trim(p4))
Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
    
If seldrequisitn_rs.RowCount > 0 Then

    txt_reqno.Text = Trim(seldrequisitn_rs("reqno"))

    SELRREQNO_QR(0) = Trim(seldrequisitn_rs("reqno"))
    Set SELRREQNO_RS = SELRREQNO_QR.OpenResultset(1, 3)

    selta101qr(0) = Trim(seldrequisitn_rs("caseno"))
    Set selta101rs = selta101qr.OpenResultset(1, 3)
        
    datetime
    
    If selta101rs.RowCount > 0 Then
        txt_name.Text = Trim(selta101rs("tag")) & " " & Trim(selta101rs("first_name")) & "" & Trim(selta101rs("Mid_name")) & " " & Trim(selta101rs("last_name"))
        txt_sex.Text = Trim(selta101rs("SEX"))
        PATYPE = Trim(selta101rs("PAT_TYPE"))
        txt_Age.Text = DateDiff("YYYY", selta101rs("BIRTH_DATE"), CStr(DATE11))
        txt_catgory.Text = Trim(selta101rs("PAT_CATG"))
        msk_caseno.Text = Trim(selta101rs("CASE_NO"))
    End If
        
    sel_rtime_qr(0) = Trim(seldrequisitn_rs("reqno"))
    Set sel_rtime_rs = sel_rtime_qr.OpenResultset(1, 3)
        
    If Not sel_rtime_rs.EOF Then
        TXT_COLLDTTM.Text = Trim(sel_rtime_rs("SAMCOLDATE")) & " / " & Trim(sel_rtime_rs("SAMCOLTIME"))
        TXT_RCVDTTM.Text = Trim(sel_rtime_rs("SAMRECDATE")) & " / " & Trim(sel_rtime_rs("SAMRECTIME"))
    End If
        
Else
    MsgBox "Path No. Does not Exists"
End If

End Sub

Private Sub txt_printQty_KeyPress(KeyAscii As Integer)
If KeyAscii > 47 And KeyAscii < 58 Then
Else
KeyAscii = 0
End If
End Sub

Private Sub txt_seqno_KeyPress(KeyAscii As Integer)

If ((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8) Or (KeyAscii = 13) Or (KeyAscii = 32) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or ((KeyAscii >= 65) And (KeyAscii <= 90)) Then
Else
'    If KeyAscii = 13 Then
'    Else
    KeyAscii = 0
    'End If
End If

Dim i As Integer
Dim seqno As String, FSTEXT As String
If KeyAscii = 13 Then

'    If UCase(Right(Trim(txt_seqno.Text), 2)) = "FS" Then
'        seqno = Left(Trim(txt_seqno.Text), Len(txt_seqno.Text) - 2)
'    Else
'        seqno = Trim(txt_seqno.Text)
'    End If



    FSTEXT = Trim(txt_seqno.Text)
    If Trim(FSTEXT) <> "" Then
      Arr = Split(FSTEXT, " ")
      seqno = Arr(0)
      'bctxt1 = ARR(1) & " " & ARR(2)
    End If

       ' seqno = Trim(Left(Trim(txt_seqno.Text), 2))

    
    If CInt(Trim(seqno)) > CInt(Trim(txt_printQty.Text)) Then
        MsgBox "Entered number is greater than Print quantity", vbCritical, "Invalid Number"
        txt_seqno.SelStart = 0
        txt_seqno.SelLength = Len(txt_seqno.Text)
        Exit Sub
    End If
    
    If CInt(Trim(seqno)) < prevseqno Then
        MsgBox "Entered Number is less than previous one", vbCritical, "ERROR"
         txt_seqno.SelStart = 0
        txt_seqno.SelLength = Len(txt_seqno.Text)
        Exit Sub
    End If
    
    If CInt(Trim(seqno)) > 99 Then
        MsgBox "Upto 99 numbers allowed"
        txt_seqno.SelStart = 0
        txt_seqno.SelLength = Len(txt_seqno.Text)
        Exit Sub
    End If
    
    
    
    Dim CHKNO As Integer
    For i = 1 To MSFlexGrid1.Rows - 1
        If IsNumeric(Left(Trim(MSFlexGrid1.TextMatrix(i, 1)), 2)) = True Then
            CHKNO = Left(Trim(MSFlexGrid1.TextMatrix(i, 1)), 2)
        Else
            If IsNumeric(Left(Trim(MSFlexGrid1.TextMatrix(i, 1)), 1)) = True Then
            CHKNO = Left(Trim(MSFlexGrid1.TextMatrix(i, 1)), 1)
            End If
        End If
        
        
            If CHKNO = Trim(seqno) Then
                MsgBox "Number already given"
                 txt_seqno.SelStart = 0
                 txt_seqno.SelLength = Len(txt_seqno.Text)
                Exit Sub
            End If
       
    Next
    
    txt_seqno.Visible = True
    
    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 0) = MSFlexGrid1.row
    
    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1) = txt_seqno.Text
    
    MSFlexGrid1.col = 1
    
    MSFlexGrid1.Rows = MSFlexGrid1.row + 2
    MSFlexGrid1.row = MSFlexGrid1.row + 1
    
    txt_seqno.Visible = True
    txt_seqno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
    prevseqno = CInt(Trim(seqno))
    txt_seqno.Text = ""
    
        
End If
End Sub

Private Sub TXTPATHG_Change()
'On Error Resume Next
If Len(Trim(TXTPATHG.Text)) = 1 And txt_pathno.Enabled = True Then
txt_pathno.SetFocus
Else

MsgBox "Please Click on View Button First!!!!!!!"
TXTPATHG.Text = ""
cmb_view.SetFocus
End If
End Sub

Private Sub TXTPATHG_KeyPress(KeyAscii As Integer)
If KeyAscii = 71 Or KeyAscii = 8 Or KeyAscii = 13 Then

Else
    If KeyAscii = 32 Or KeyAscii = 103 Then
    KeyAscii = 71
    Else
    KeyAscii = 0
    End If

End If
End Sub
