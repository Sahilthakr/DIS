VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_HeamatBarcode 
   BackColor       =   &H8000000B&
   Caption         =   "HeamatPathology  Barcode Receiving "
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11070
   LinkTopic       =   "Form3"
   ScaleHeight     =   9045
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_ok 
      BackColor       =   &H80000003&
      Caption         =   "&Save"
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
      Left            =   3480
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   8040
      Width           =   855
   End
   Begin VB.TextBox TXT_SLNO 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3360
      TabIndex        =   44
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox TXT_BEDNO 
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
      Left            =   5880
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox TXT_WARD 
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
      Left            =   2640
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   840
      TabIndex        =   42
      Top             =   840
      Width           =   2295
      Begin VB.OptionButton Opt_patient 
         Caption         =   "Patient"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton opt_staff 
         Caption         =   "Staff"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox Text1 
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
      Left            =   1440
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2760
      Width           =   615
   End
   Begin VB.ComboBox txt_reqno 
      Height          =   315
      Left            =   5880
      TabIndex        =   3
      Text            =   "Combo2"
      Top             =   1560
      Width           =   3015
   End
   Begin VB.ComboBox CMB_StainType 
      Height          =   315
      ItemData        =   "frm_HeamatBarcode.frx":0000
      Left            =   3960
      List            =   "frm_HeamatBarcode.frx":0002
      TabIndex        =   37
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox txt_catEgory 
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
      Left            =   4320
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2400
      Width           =   735
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
      Left            =   2640
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2400
      Width           =   615
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
      Left            =   1440
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2400
      Width           =   375
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
      Left            =   1440
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2040
      Width           =   7455
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
      Left            =   2040
      TabIndex        =   12
      Top             =   3600
      Width           =   3495
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
      Left            =   7320
      TabIndex        =   13
      Top             =   4200
      Width           =   3495
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   975
      Left            =   1080
      TabIndex        =   16
      Top             =   7680
      Width           =   7815
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000003&
         Caption         =   "&New"
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
         Left            =   480
         MaskColor       =   &H0080C0FF&
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmd_back 
         BackColor       =   &H80000003&
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
         Left            =   4440
         MaskColor       =   &H0080C0FF&
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmb_print 
         BackColor       =   &H80000003&
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
         Left            =   3360
         MaskColor       =   &H0080C0FF&
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmd_cancel 
         BackColor       =   &H80000003&
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
         Left            =   5520
         MaskColor       =   &H0080C0FF&
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmb_view 
         BackColor       =   &H80000003&
         Caption         =   "&View"
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
         Left            =   1440
         MaskColor       =   &H0080C0FF&
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmd_exit 
         BackColor       =   &H80000003&
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
         TabIndex        =   17
         Top             =   360
         Width           =   855
      End
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
      Left            =   2040
      TabIndex        =   14
      Top             =   4200
      Width           =   3495
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
      Left            =   2040
      TabIndex        =   4
      Text            =   "1"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox txt_ipno 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2040
      MaxLength       =   14
      TabIndex        =   15
      Top             =   4800
      Width           =   3495
   End
   Begin MSMask.MaskEdBox msk_caseno 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      ToolTipText     =   "ENTER THE CASE NUMBER TO DISPLAY PATIENT DETAILS"
      Top             =   1560
      Width           =   2415
      _ExtentX        =   4260
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
      Left            =   3240
      TabIndex        =   22
      Top             =   5280
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   3625
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   14737632
      ForeColorFixed  =   128
      GridColor       =   0
      ScrollBars      =   0
      AllowUserResizing=   3
      Appearance      =   0
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
   Begin VB.Label lbl_bed 
      BackStyle       =   0  'Transparent
      Caption         =   "BedNo"
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
      Left            =   5160
      TabIndex        =   43
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
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
      Left            =   5160
      TabIndex        =   40
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label lbl_ward 
      BackStyle       =   0  'Transparent
      Caption         =   "Ward"
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
      Left            =   2160
      TabIndex        =   39
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Pat_Status"
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
      TabIndex        =   38
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "IP No."
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
      Left            =   720
      TabIndex        =   36
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "HEAMATO PATHOLOGY BARCODE LABEL PRINT"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2400
      TabIndex        =   35
      Top             =   480
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Heamatology No."
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
      Left            =   120
      TabIndex        =   34
      Top             =   4200
      Width           =   1695
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
      Left            =   5280
      TabIndex        =   33
      Top             =   840
      Width           =   1695
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
      Left            =   600
      TabIndex        =   32
      Top             =   1680
      Width           =   855
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
      Left            =   3360
      TabIndex        =   31
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label lbl_sex 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
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
      Left            =   720
      TabIndex        =   30
      Top             =   2400
      Width           =   675
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
      Left            =   720
      TabIndex        =   29
      Top             =   2040
      Width           =   525
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
      Left            =   2160
      TabIndex        =   28
      Top             =   2400
      Width           =   360
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
      Left            =   120
      TabIndex        =   27
      Top             =   3600
      Width           =   1815
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
      Left            =   5640
      TabIndex        =   26
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
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
      Height          =   375
      Left            =   240
      TabIndex        =   25
      Top             =   0
      Width           =   9495
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
      Left            =   4920
      TabIndex        =   24
      Top             =   1560
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
      Left            =   360
      TabIndex        =   23
      Top             =   5280
      Width           =   1455
   End
End
Attribute VB_Name = "frm_HeamatBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim slno1 As Integer
Dim selsclqr As rdoQuery
Dim selsclrs As rdoResultset

Dim sel_rtime_qr As rdoQuery
Dim sel_rtime_rs As rdoResultset

Dim req_qr As rdoQuery
Dim req_rs As rdoResultset

Dim seldrequisitn_qr As rdoQuery
Dim seldrequisitn_rs As rdoResultset

Dim SELTA101_QR As rdoQuery
Dim SELTA101_RS As rdoResultset


Dim selma104_qr As rdoQuery
Dim selma104_rs As rdoResultset


Dim seldrequisitn_qrC As rdoQuery
Dim seldrequisitn_rsC As rdoResultset

Dim SELADDENDUMCYQR As rdoQuery
Dim SELADDENDUMCYRS As rdoResultset

Dim SELRREQNO_QR As rdoQuery
Dim SELRREQNO_RS As rdoResultset



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
Dim NEWFLAG As Boolean, REQFLAG1 As Boolean

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
'On Error Resume Next

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
        INS_PATHBRCLBL_QR(0) = Replace(Trim(txt_reqno.Text), "/", "")
        INS_PATHBRCLBL_QR(1) = Trim(msk_caseno.FormattedText)
        INS_PATHBRCLBL_QR(2) = txt_ipno.Text
        INS_PATHBRCLBL_QR(3) = ""
        INS_PATHBRCLBL_QR(4) = Trim(txt_category.Text)
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


'Dim printreason As String
'Dim lprng As Integer
'If chk_slide.Value = 1 And chk_block.Value = 1 Then
'printreason = "12"
'lprng = 2
'Else
'    If chk_slide.Value = 1 Then
'    printreason = "1"
'    lprng = 1
'    End If
'    If chk_block = 1 Then
'    printreason = "2"
'    lprng = 1
'    End If
'End If


Dim FSTEXT As String
Dim pq, LP, l, r, ln1 As Integer
Dim D1, rm1, bctxt, bctxt1 As String
pq = txt_printQty.Text
If pq > 0 Then
'SAVEDATA
f1 = FreeFile()
Open "c:\DISREPORTSEC\DATA1\micro.TXT" For Output As #f1
For LP = 1 To pq
'
'{D0231,0225,0201|}
'{AY;+10,0|}
'{C|}
'{PV00;0008,0021,0020,0020,B,00,B=Mr.Test Test Test|}
'{XB00;0035,0064,B,1,01,01,03,03,01,0,0032,+0000000000,0,00=12345678|}
'{PV01;0046,0121,0021,0021,B,00,B=12345678|}
'{PV02;0114,0047,0020,0020,B,00,B=RH/12345|}
'{PV03;0008,0147,0020,0020,B,00,B=IP00012|}
'{PV04;0131,0141,0020,0020,B,00,B=1 MPO|}
'{XS;I,0001,0002C4200|}

'' commented following by shweta 20-12-2013 'worked previously
PRNTSTR1 = ""
'PRNTSTR1 = "TEC BARCODE PRINTER JOB START" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{D0231,0225,0201|}" & vbCrLf & "{AY;+10,0|}" & vbCrLf '{D0221,0225,0201|}
'PRNTSTR1 = PRNTSTR1 & "{C|}" & vbCrLf & "{PV00;0008,0021,0020,0020,B,00,B=" & Trim(txt_name.Text) & "|}" & vbCrLf
'''till here
'PRNTSTR1 = PRNTSTR1 & "{XB00;0019,0050,A,3,02,0,0047,+0000000000,000,0,00|}" '& Trim(msk_caseno.FormattedText) & "|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{PC000;0080,0040,05,05,J,00,B|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{LC;0080,0041,0167,0041,0,1|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{PC001;0032,0174,1,1,C,00,B|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{PC002;0102,0174,05,1,R,00,B|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{PC003;0047,0019,05,05,J,00,B|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{LC;0047,0021,0199,0021,0,1|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{PC004;0018,0136,1,1,D,00,B|}" & vbCrLf

'PRNTSTR1 = "TEC BARCODE PRINTER JOB START" & vbCrLf
'----changed bymrunal
'changes by shweta 25-09-2014
'PRNTSTR1 = PRNTSTR1 & "{D0245,0200,0225|}" & vbCrLf & "{AY;+10,0|}" & vbCrLf '{D0221,0225,0201|}
PRNTSTR1 = PRNTSTR1 & "{D0200,0480,0180|}" & vbCrLf & "{AY;+10,0|}" & vbCrLf '{D0221,0225,0201|}
'PRNTSTR1 = PRNTSTR1 & "{C|}" & vbCrLf & "{PV00;0015,0025,0024,0028,B,00,B=" & Trim(Msk_caseno.Text) & "|}" & vbCrLf
PRNTSTR1 = PRNTSTR1 & "{C|}" & vbCrLf & "{PV00;0032,0011,0025,0020,B,00,B=" & Trim(msk_caseno.Text) & "|}" & vbCrLf
'-----
bctxt1 = ""
 bctxt = LP
' ln1 = bctxt
 Dim CHKNO As Integer
 Dim PRINTNO As String
 Dim PRINTDESC As String
''''''''''''
PRINTNO = LP
   PRINTDESC = ""
   If MSFlexGrid1.TextMatrix(1, 0) = "" Then
   Else
 For i = 1 To MSFlexGrid1.Rows - 1
 On Error Resume Next
   If MSFlexGrid1.TextMatrix(i, 0) = "" Then
   PRINTNO = i
   PRINTDESC = ""
   Else
   PRINTNO = MSFlexGrid1.TextMatrix(i, 0)
   PRINTDESC = MSFlexGrid1.TextMatrix(i, 1)
   End If
'
'        If IsNumeric(Left(Trim(MSFlexGrid1.TextMatrix(i, 1)), 2)) = True Then
'            ln1 = Left(Trim(MSFlexGrid1.TextMatrix(i, 1)), 2)
'        Else
'            If IsNumeric(Left(Trim(MSFlexGrid1.TextMatrix(i, 1)), 1)) = True Then
'            ln1 = Left(Trim(MSFlexGrid1.TextMatrix(i, 1)), 1)
'            End If
'        End If

If ln1 = LP Then
D1 = PRINTDESC 'Trim(MSFlexGrid1.TextMatrix(i, 1))

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
 End If
 'commented by shweta 30-jul-2014 as suggested by kalpana madam to print ipno for barcode
'PRNTSTR1 = PRNTSTR1 & "{XB00;0016,0047,B,1,01,01,03,03,01,0,0032,+0000000000,0,00=" & Trim(txt_pathno.Text) & "|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{PV01;0004,0110,0035,0035,B,00,B=" & Trim(txt_pathno.Text) & "|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{PV02;0016,0134,0024,0024,B,00,B=" & Trim(txt_ipno.Text) & "|}" & vbCrLf
''PRNTSTR1 = PRNTSTR1 & "{PV03;0159,0137,0028,0028,B,00,B=" & Trim(PRINTNO) & " " & Trim(PRINTDESC) & "|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{PV03;0159,0137,0028,0028,B,00,B=" & "" & " " & "" & "|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{XS;I,0001,0002C4200|}" & vbCrLf
 ''till here
 ''and follpwing added
 
 PRNTSTR1 = PRNTSTR1 & "{XB00;0000,0016,B,1,02,02,05,05,02,0,0064,+0000000000,0,00=" & Trim(txt_pathno.Text) & "|}" & vbCrLf
PRNTSTR1 = PRNTSTR1 & "{PV01;0004,0110,0035,0035,B,00,B=" & Trim(txt_pathno.Text) & "|}" & vbCrLf
PRNTSTR1 = PRNTSTR1 & "{PV02;0016,0134,0024,0024,B,00,B=" & Trim(txt_ipno.Text) & "|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{PV03;0159,0137,0028,0028,B,00,B=" & Trim(PRINTNO) & " " & Trim(PRINTDESC) & "|}" & vbCrLf
PRNTSTR1 = PRNTSTR1 & "{PV03;0159,0137,0028,0028,B,00,B=" & "" & " " & "" & "|}" & vbCrLf
PRNTSTR1 = PRNTSTR1 & "{XS;I,0001,0002C4200|}" & vbCrLf
'----removed by mrunal
' PRNTSTR1 = PRNTSTR1 & "{XB00;0016,0047,B,1,01,01,03,03,01,0,0032,+0000000000,0,00=" & Trim(txt_pathno.Text) & "|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{PV01;0004,0110,0035,0035,B,00,B=" & Trim(txt_pathno.Text) & "|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{PV02;0016,0134,0024,0024,B,00,B=" & Trim(txt_ipno.Text) & "|}" & vbCrLf
''PRNTSTR1 = PRNTSTR1 & "{PV03;0159,0137,0028,0028,B,00,B=" & Trim(PRINTNO) & " " & Trim(PRINTDESC) & "|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{PV03;0159,0137,0028,0028,B,00,B=" & "" & " " & "" & "|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{XS;I,0001,0002C4200|}" & vbCrLf
'----
'commented by shweta 20-12-2013 'worked previously
'PRNTSTR1 = PRNTSTR1 & "{XB00;0035,0064,B,1,01,01,03,03,01,0,0032,+0000000000,0,00=" & Trim(txt_pathno.Text) & "|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{PV01;0044,0121,0021,0021,B,00,B=" & Trim(txt_pathno.Text) & "|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{PV02;0114,0047,0020,0020,B,00,B=" & Trim(msk_caseno.FormattedText) & "|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{PV03;0008,0147,0020,0020,B,00,B=" & Trim(txt_ipno.Text) & "|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{PV04;0131,0141,0020,0020,B,00,B=" & Trim(PRINTNO) & " " & Trim(PRINTDESC) & "|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{XS;I,0001,0002C4200|}" & vbCrLf

''till here 20-12-2013 'worked previously

'PRNTSTR1 = PRNTSTR1 & "{RB00;>6G>5" & Replace(Left(Trim(txt_pathno.Text), Len(txt_pathno.Text) - 2), "/", "") & ">6" & Trim(Right(Trim(txt_pathno.Text), 2)) & "|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{RC004;" & Trim(txt_pathno.Text) & "|}" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "{PC004;0005,0136,1,1,D,00,B=" & Replace(Trim(txt_pathno.Text), "/", "") & "|}" & vbCrLf

'PRNTSTR1 = PRNTSTR1 & "{XS;I,0001,0002C4200|} " & vbCrLf
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



End If


Command1.Enabled = False
cmb_view.Enabled = False
cmd_ok.Enabled = False
cmb_print.Enabled = True
cmd_cancel.Enabled = True
cmd_back.Enabled = True

Exit Sub
AA:
rdb.RollbackTrans

End Sub





Private Sub CMB_StainType_Click()
If MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 0) = "" Then
CMB_StainType.Text = ""
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1) = ""
Exit Sub
End If
If (Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 0)) + 1) > txt_printQty.Text Then

Exit Sub
End If
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1) = CMB_StainType.Text
MSFlexGrid1.Rows = MSFlexGrid1.row + 2
    MSFlexGrid1.row = MSFlexGrid1.row + 1
    TXT_SLNO.Visible = True
    MSFlexGrid1.col = 0
        MSFlexGrid1_Click
    CMB_StainType.Visible = True
    CMB_StainType.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
    'prevseqno = CInt(Trim(seqno))
    CMB_StainType.Text = ""
    TXT_SLNO.SetFocus
End Sub

Private Sub cmd_Ok_Click()
Dim prpathno As String
Dim entr, CytoFlag As String
'On Error GoTo MSG:
'' setting of tabs based on dhemcyto1 table
'      SSTab1.TabEnabled(0) = False
'      SSTab1.TabEnabled(1) = False
'      SSTab1.TabEnabled(2) = False
'      SSTab1.TabEnabled(3) = False
'      SSTab1.TabEnabled(4) = False

'************ Shraddha Pandit 30/03/09 ****************
'******************************************************



 '' seting of tabs based on dhemcyto1 table over
If Opt_patient.Value = True Or Opt_patient.Value = False Then
        If NEWFLAG = True Then
            NEWFLAG = False
            If selta101flag1 = False Then
                SELTA101_QR(0) = msk_caseno.FormattedText
                Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
                selta101flag1 = True
            Else
                SELTA101_QR(0) = msk_caseno.FormattedText
                SELTA101_RS.MoreResults
                SELTA101_RS.Requery
            End If
            If SELTA101_RS.RowCount = 0 Then
                Exit Sub
            End If
        datetime
            txt_sex.Text = SELTA101_RS("sex")
            txt_name.Text = Trim(SELTA101_RS("first_name")) & " " & Trim(SELTA101_RS("mid_name")) & " " & Trim(SELTA101_RS("last_name"))
            If Len(Trim(SELTA101_RS("BIRTH_DATE"))) <> 0 Then
                    'PATYPE = Trim(SELTA101_RS("PAT_TYPE"))
                   YEAR1 = DateDiff("m", SELTA101_RS("BIRTH_DATE"), DATE11)

                   Dim n As Integer
                   n = YEAR1 / 12
                   txt_age.Text = n
               End If

               If Trim(txt_age.Text) = 0 Then
                  Dim age As String
                  age = DateDiff("m", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                  age = "." & (age Mod 12)
                     If age = ".0" Then
                        age = DateDiff("d", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                        age = ".0" & (age) ' Mod 30)
                     End If
                  txt_age.Text = Trim(age)
               End If
            'txt_age = DateDiff("YYYY", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
          '  txt_pridoc.Text = SELTA101_RS("ud_name")
            txt_category.Text = SELTA101_RS("pat_catg")


            If SELMA104FLAG1 = False Then
                selma104_qr(0) = msk_caseno.FormattedText
                Set selma104_rs = selma104_qr.OpenResultset(1, 3)
                SELMA104FLAG1 = True
            Else
                selma104_qr(0) = msk_caseno.FormattedText
                selma104_rs.MoreResults
                selma104_rs.Requery
            End If

            If selma104_rs.RowCount > 0 Then
                txt_ward.Visible = True
                txt_bedno.Visible = True
                lbl_ward.Visible = True
                lbl_bed.Visible = True
                txt_ward.Text = selma104_rs("WARD_NAME")
                txt_bedno.Text = selma104_rs("bed_no")
                Text1.Text = "IP"
            Else
                Text1.Text = "OP"
                txt_ward.Visible = False
                txt_bedno.Visible = False
                lbl_ward.Visible = False
                lbl_bed.Visible = False
            End If




''
End If
End If

sel_rtime_qr(0) = Trim(Replace(txt_reqno.Text, "/", ""))
    Set sel_rtime_rs = sel_rtime_qr.OpenResultset(1, 3)

    If Not sel_rtime_rs.EOF Then
        TXT_COLLDTTM.Text = Trim(sel_rtime_rs("SAMCOLDATE")) & " / " & Trim(sel_rtime_rs("SAMCOLTIME"))
        TXT_RCVDTTM.Text = DATE11 & " / " & time1 '  Trim(sel_rtime_rs("SAMRECDATE")) & " / " & Trim(sel_rtime_rs("SAMRECTIME"))
    End If
'''receiving Number

'receiveNo (Trim(Replace(txt_reqno.Text, "/", "")))

'--------21-11-2013 increament path no on save button----------
SPAUTONO
IPNO


'''''
Dim UPDRTIME_QR As rdoQuery
Dim UPDRREQ_QR As rdoQuery
Dim UPDDREQ_QR As rdoQuery
'On Error GoTo AA
rdb.BeginTrans
str = "UPDATE RISLIB.RTIMELOG SET SAMRECDATE=current date, SAMRECTIME=current time WHERE REQNNO=? "
Set UPDRTIME_QR = rdb.CreateQuery("", str)
UPDRTIME_QR(0) = Replace(Trim(txt_reqno.Text), "/", "")
UPDRTIME_QR.Execute
str = "UPDATE RISLIB.RREQUISITN SET REQNSTAT='E' WHERE REQNNO=?"
Set UPDRREQ_QR = rdb.CreateQuery("", str)
UPDRREQ_QR(0) = Replace(Trim(txt_reqno.Text), "/", "")
UPDRREQ_QR.Execute

str = "UPDATE DISLIB.DREQUISITN SET PATHNO=? WHERE REQNO=?"
Set UPDDREQ_QR = rdb.CreateQuery("", str)
UPDDREQ_QR(0) = Trim(txt_pathno.Text)
UPDDREQ_QR(1) = Replace(Trim(txt_reqno.Text), "/", "")
UPDDREQ_QR.Execute
rdb.CommitTrans


'''''


SAVEDATA

Command1.Enabled = False
cmb_view.Enabled = False
cmd_ok.Enabled = False
cmb_print.Enabled = True
cmd_cancel.Enabled = True
cmd_back.Enabled = True
Exit Sub
AA:

rdb.RollbackTrans
End Sub

Private Sub Command1_Click()
NEWFLAG = True
cmd_cancel.Enabled = True
cmb_print.Enabled = True
cmd_back.Enabled = True
cmb_view.Enabled = False
cmd_ok.Enabled = True
txt_pathno.Enabled = True
txt_pathno.Text = ""
txt_name.Text = ""
msk_caseno.Enabled = True
msk_caseno.SetFocus
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
str = "SELECT * FROM RISLIB.RTIMELOG WHERE REQNNO=?"
Set sel_rtime_qr = rdb.CreateQuery("", str)

str = "SELECT * FROM PATOOLSLIB.TA101 WHERE CASE_NO=?"
Set selta101qr = rdb.CreateQuery("", str)

str = "SELECT * FROM PATOOLSLIB.TA101 WHERE CASE_NO=?"
Set SELTA101_QR = rdb.CreateQuery("", str)


str = "select * from patoolslib.ta1012 where case_no=?"
Set selta1012_qr = rdb.CreateQuery("", str)

str = "SELECT * FROM PATOOLSLIB.MA301 WHERE SITECD=?"
Set SELMA301_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO=?"
Set SELDREQNOQR = rdb.CreateQuery("", str)

   str = "select * from PATOOLSLIB.ma104 where case_no =?"
    Set selma104_qr = rdb.CreateQuery("selma104_qr11HE", str)


str = "INSERT INTO DISLIB.PATHBRCLBL(REQNO,CASENO,PATHNO,COLLUNIT,PAT_CATG,PRNTQTY,SEQNO) VALUES(?,?,?,?,?,?,?)"
Set INS_PATHBRCLBL_QR = rdb.CreateQuery("", str)

Set req_qr = rdb.CreateQuery("reqBM", "select * from Rislib.RREQUISITN where caseno=? AND (REQNSTAT='C' OR REQNSTAT='P') AND dept_code='UZZ' ORDER BY REQNDT desc ")



str = "SELECT DESCRDET FROM DISLIB.MHEMRECEIV WHERE GROUP='STAIN' "
Set selsclqr = rdb.CreateQuery("", str)
Set selsclrs = selsclqr.OpenResultset(1, 3)
If selsclrs.RowCount > 0 Then
selsclrs.MoveFirst
Dim KL As Integer
    For KL = 0 To selsclrs.RowCount - 1
        CMB_StainType.AddItem (Trim(selsclrs("DESCRDET")))
        selsclrs.MoveNext
    Next
Else
MsgBox "DATA NOT FOUND IN STAIN MASTER", vbExclamation
End If


'txt_pathno.Text = ""
'txt_name.Text = ""

txt_pathno.Enabled = False

Command1.Enabled = True
cmb_view.Enabled = True
cmd_ok.Enabled = False
cmb_print.Enabled = False
cmd_cancel.Enabled = True
cmd_back.Enabled = True


End Sub


Private Sub cmb_view_Click()
cmd_cancel.Enabled = True
cmb_print.Enabled = True
cmd_back.Enabled = True
cmb_view.Enabled = False
cmd_ok.Enabled = False


txt_pathno.Enabled = True
txt_pathno.Text = ""
txt_name.Text = ""
'TXTPATHG.Text = ""
'txt_pathno.SetFocus
'TXTPATHG.SetFocus
End Sub

Private Sub cmd_cancel_Click()
prevseqno = 0
Command1.Enabled = True
cmb_view.Enabled = True
cmd_ok.Enabled = False
cmb_print.Enabled = False
cmd_cancel.Enabled = True
cmd_back.Enabled = True



txt_pathno.Enabled = False

MSFlexGrid1.row = 1
MSFlexGrid1.col = 1
'txt_seqno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top

'txt_seqno.Visible = False
MSFlexGrid1.clear

txt_pathno.Text = ""
txt_name.Text = ""
msk_caseno.Text = ""
txt_age.Text = ""
txt_sex.Text = ""
txt_category.Text = ""
TXT_COLLDTTM.Text = ""
TXT_RCVDTTM.Text = ""
txt_printQty.Text = ""
txt_reqno.Text = ""
'TXTPATHG.Text = ""
txt_ipno.Text = ""
prevseqno = 0
cmb_view.Enabled = True
TXT_SLNO.Text = ""
TXT_SLNO.Visible = False
CMB_StainType.Text = ""
CMB_StainType.Visible = False
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








Private Sub MSFlexGrid1_Click()
If MSFlexGrid1.col = 0 Then

        TXT_SLNO.Visible = True
        TXT_SLNO.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
        TXT_SLNO.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
        TXT_SLNO.Width = MSFlexGrid1.CellWidth + 10
       ' TXT_SLNO.Text = MSK_MANUALREC.TextMatrix(MSK_MANUALREC.row, 4)
      '  TXT_SLNO.SelStart = 6
       ' TXT_SLNO.SelLength = 4
       TXT_SLNO.Text = slno1 + 1
       TXT_SLNO.SetFocus
End If
End Sub

Private Sub msk_caseno_LostFocus()
Dim CASENO As String
On Error GoTo MSG:

'If Len(Trim(MaskEdBox1.Text)) = "" Then
        If Trim(msk_caseno.Text) = "" Then 'commented & added by Shraddha Pandit on 06/04/09
Exit Sub
End If
        CASENO = msk_caseno.Text
       '' msk_caseno.Text = UCase(padzero(5, "C", CASENO, 3)) commented by Monika on 05/12/2022
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
If Len(msk_caseno.FormattedText) = 6 Then
Else

         If NEWFLAG = True Then
            If REQFLAG1 = False Then
                req_qr(0) = msk_caseno.FormattedText
               ' req_qr(1) = Trim(txt_modcode.Text)
                Set req_rs = req_qr.OpenResultset(1, 3)
            Else
                req_qr(0) = msk_caseno.FormattedText
                'req_qr(1) = Trim(txt_modcode.Text)
                req_rs.MoreResults
                req_rs.Requery
          End If
          If req_rs.RowCount > 0 Then
                txt_reqno.clear
                While Not req_rs.EOF

                    txt_reqno.AddItem req_chglen(req_rs("reqNno"))
                    req_rs.MoveNext

                Wend
                    'cmb_reqno.ListIndex = 0 'Commented by Shraddha on 10/09/10
          Else
                txt_reqno.clear
          End If
       End If
'          If SELDREQNOFLAG = False Then
'            SELDREQNOQR(0) = MaskEdBox1.FormattedText
'            Set SELDREQNORS = SELDREQNOQR.OpenResultset(1, 3)
'            SELDREQNOFLAG = True
'          Else
'            SELDREQNOQR(0) = MaskEdBox1.FormattedText
'            SELDREQNORS.MoreResults
'            SELDREQNORS.Requery
'          End If
'          If SELDREQNORS.RowCount > 0 Then
'            CMB_PATHNO.clear
'            While Not SELDREQNORS.EOF
'                CMB_PATHNO.AddItem SELDREQNORS("PATHNO")
'                SELDREQNORS.MoveNext
'            Wend
'          End If
        End If

'        If viewflag = True Then
'
'            If SELRREQUISITNFLAG1 = False Then
'                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
'                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
'                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
'                SELRREQUISITN_QR(3) = "T"
'                SELRREQUISITN_QR(4) = "F"
'                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
'                SELRREQUISITNFLAG1 = True
'            Else
'                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
'                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
'                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
'                SELRREQUISITN_QR(3) = "T"
'                SELRREQUISITN_QR(4) = "F"
'                SELRREQUISITN_RS.MoreResults
'                SELRREQUISITN_RS.Requery
'            End If
'            If SELRREQUISITN_RS.RowCount > 0 Then
'                cmb_reqno.clear
'                While Not SELRREQUISITN_RS.EOF
'                    cmb_reqno.AddItem req_chglen(SELRREQUISITN_RS("REQNNO"))
'                    SELRREQUISITN_RS.MoveNext
'                Wend
'             Else
'                MsgBox " No  Requisition for this Modality"
'
'                btnenabl ("INI")
'                Exit Sub
'            End If
'
'        End If
'End If
'    If opt_staff.Value = True Then
'        CASENO = MaskEdBox1.Text
'        If PREPATH_FLAG = False Then
'            PREPATH_QR(0) = Trim(MaskEdBox1.FormattedText)
'            Set PREPATH_RS = PREPATH_QR.OpenResultset(1, 3)
'            PREPATH_FLAG = True
'        Else
'            PREPATH_QR(0) = Trim(MaskEdBox1.FormattedText)
'            PREPATH_RS.MoreResults
'            PREPATH_RS.Requery
'        End If
'        If Not PREPATH_RS.EOF Then
'            Do While Not PREPATH_RS.EOF
'                CMB_PREPATHNO.AddItem req_chglen(PREPATH_RS("REQNNO")) & " " & "$" & PREPATH_RS("PATHNO")
'                PREPATH_RS.MoveNext
'            Loop
'        End If
'        'MaskEdBox1.Text = UCase(padzero(5, "C", CASENO, 3))
'         If NEWFLAG = True Then
'            If REQFLAG1 = False Then
'                req_qr(0) = Trim(MaskEdBox1.FormattedText)
'                req_qr(1) = MODALITY1
'                Set req_rs = req_qr.OpenResultset(1, 3)
'            Else
'                req_qr(0) = Trim(MaskEdBox1.FormattedText)
'                req_qr(1) = MODALITY1
'                req_rs.MoreResults
'                req_rs.Requery
'          End If
'          If req_rs.RowCount > 0 Then
'                cmb_reqno.clear
'                While Not req_rs.EOF
'                cmb_reqno.AddItem req_chglen(req_rs("reqNno"))
'
'                req_rs.MoveNext
'
'                Wend
'          Else
'            cmb_reqno.clear
'
'            End If
'
'
'        End If
'
'        If viewflag = True Then
'
'            If SELRREQUISITNFLAG1 = False Then
'                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
'                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
'                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
'                SELRREQUISITN_QR(3) = "T"
'                SELRREQUISITN_QR(4) = "F"
'                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
'                SELRREQUISITNFLAG1 = True
'            Else
'                SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
'                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
'                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
'                SELRREQUISITN_QR(3) = "T"
'                SELRREQUISITN_QR(4) = "F"
'                SELRREQUISITN_RS.MoreResults
'                SELRREQUISITN_RS.Requery
'            End If
'            If SELRREQUISITN_RS.RowCount > 0 Then
'                cmb_reqno.clear
'                While Not SELRREQUISITN_RS.EOF
'                    cmb_reqno.AddItem req_chglen(SELRREQUISITN_RS("REQNNO"))
'                    SELRREQUISITN_RS.MoveNext
'                Wend
'             Else
'                MsgBox " No  Requisition for this Modality"
'                btnenabl ("INI")
'                Exit Sub
'            End If
'
'        End If
'End If
Command1.Enabled = False
cmb_view.Enabled = False
cmd_ok.Enabled = True
cmb_print.Enabled = False
cmd_cancel.Enabled = True
cmd_back.Enabled = True
 Exit Sub
MSG:
MsgBox Error(Err)

End Sub

Private Sub txt_pathno_LostFocus()
'txt_printQty.Text = ""
'txt_reqno.Text = ""
If Trim(txt_pathno.Text) = "" Then
    'MsgBox "Path No. Should Not Be Empty"
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
CMB_StainType.Visible = True
MSFlexGrid1.row = 1
MSFlexGrid1.col = 1
CMB_StainType.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
CMB_StainType.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top

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
        txt_age.Text = DateDiff("YYYY", selta101rs("BIRTH_DATE"), CStr(DATE11))
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

 Function SPAUTONO() As String
 
Dim INSDAUTOPATH_QR As rdoQuery
str = "INSERT INTO DISLIB.DAUTOPATH (DEPT_CODE,MODALITYCD,PATHYEAR,PATHYEAR1,PATHNO) VALUES (?,?,?,?,?)"
Set INSDAUTOPATH_QR = rdb.CreateQuery("", str)
Dim PATH As String
Dim MODALITY As String
MODALITY = Mid(Trim(Replace(txt_reqno.Text, "/", "")), 4, 2)
If MODALITY = "" Then
Exit Function
End If
'On Error GoTo MSG:
DEPT1 = "UZZ"
Dim SELDAUTOPATH_QR As rdoQuery
Dim SELDAUTOPATH_RS As rdoResultset
Dim SELDAUTOPATHFLAG As Boolean

' ***** AUTOGENERATION OF THE PATHNO.
str = "SELECT * FROM DISLIB.DAUTOPATH WHERE DEPT_CODE=? AND MODALITYCD=? AND PATHYEAR1=?"
Set SELDAUTOPATH_QR = rdb.CreateQuery("", str)

Dim qr_seldmulproce_U753 As rdoQuery
Dim RS_seldmulproce_U753 As rdoResultset
str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=? AND (PRO_CODE='U753A' OR PRO_CODE='U754A')"
Set qr_seldmulproce_U753 = rdb.CreateQuery("", str)

Dim UPDDAUTOPATH_QR As rdoQuery

str = "UPDATE DISLIB.DAUTOPATH SET PATHNO=? WHERE DEPT_CODE=? AND MODALITYCD=? AND PATHYEAR1=?"
Set UPDDAUTOPATH_QR = rdb.CreateQuery("", str)


Dim DPATHYEAR_QR As rdoQuery
Dim DPATHYEAR_RS As rdoResultset
Dim DPATHYEARFLAG As Boolean


str = "SELECT * FROM DISLIB.DPATHYEAR WHERE AYEAR=?"
Set DPATHYEAR_QR = rdb.CreateQuery("", str)


i = 1

           SELDAUTOPATH_QR(2) = Trim(Right(DATE11, 2))
           qr_seldmulproce_U753(0) = Replace(Trim(txt_reqno.Text), "/", "")
            Set RS_seldmulproce_U753 = qr_seldmulproce_U753.OpenResultset(2, 3)


        If SELDAUTOPATHFLAG = False Then
            SELDAUTOPATH_QR(0) = Trim(DEPT1)
            If RS_seldmulproce_U753.EOF = False Then
            SELDAUTOPATH_QR(1) = "SM"
            Else
            SELDAUTOPATH_QR(1) = Trim(MODALITY)
            End If
            SELDAUTOPATH_QR(2) = Trim(Right(DATE11, 2))
            Set SELDAUTOPATH_RS = SELDAUTOPATH_QR.OpenResultset(2, 3)
            SELDAUTOPATHFLAG = True
        Else
            SELDAUTOPATH_QR(0) = Trim(DEPT1)
            'SELDAUTOPATH_QR(1) = Trim(MODALITY)
            If RS_seldmulproce_U753.EOF = False Then
            SELDAUTOPATH_QR(1) = "SM"
            Else
            SELDAUTOPATH_QR(1) = Trim(MODALITY)
            End If
            SELDAUTOPATH_QR(2) = Trim(Right(DATE11, 2))
            SELDAUTOPATH_RS.MoreResults
            SELDAUTOPATH_RS.Requery
        End If
        If SELDAUTOPATH_RS.RowCount > 0 Then
            AUTOPATHNO = SELDAUTOPATH_RS("PATHNO") + 1
            UPDDAUTOPATH_QR(0) = AUTOPATHNO
            UPDDAUTOPATH_QR(1) = Trim(DEPT1)
            'UPDDAUTOPATH_QR(2) = Trim(MODALITY)
            If RS_seldmulproce_U753.EOF = False Then
            UPDDAUTOPATH_QR(2) = "SM"
            Else
            UPDDAUTOPATH_QR(2) = Trim(MODALITY)
            End If
            UPDDAUTOPATH_QR(3) = Trim(Right(DATE11, 2))
            UPDDAUTOPATH_QR.Execute
        Else
            AUTOPATHNO = 1
            INSDAUTOPATH_QR(0) = Trim(DEPT1)
            'INSDAUTOPATH_QR(1) = Trim(MODALITY)
            If RS_seldmulproce_U753.EOF = False Then
            INSDAUTOPATH_QR(1) = "SM"
            Else
            INSDAUTOPATH_QR(1) = Trim(MODALITY)
            End If
            INSDAUTOPATH_QR(2) = Trim(Left(msk_caseno.FormattedText, 2))
            INSDAUTOPATH_QR(3) = Trim(Right(DATE11, 2))
            INSDAUTOPATH_QR(4) = 1
            INSDAUTOPATH_QR.Execute
        End If
        If Trim(DEPT1) = "UZZ" Then
                PATH = Trim(SELDAUTOPATH_RS("PATHYEAR"))
        Else
            If DPATHYEARFLAG = False Then
                DPATHYEAR_QR(0) = "A" & Right(DATE11, 4)
                Set DPATHYEAR_RS = DPATHYEAR_QR.OpenResultset(1, 3)
                DPATHYEARFLAG = True
            Else
                DPATHYEAR_QR(0) = "A" & Right(DATE11, 4)
                DPATHYEAR_RS.MoreResults
                DPATHYEAR_RS.Requery
            End If
            If DPATHYEAR_RS.RowCount > 0 Then
                PATH = DPATHYEAR_RS("LASTTWO")
            End If
        End If
    '***********padding the reqslno with leading zeros"****************
        Dim reqslno1
        reqslno1 = AUTOPATHNO
        If Trim(DEPT1) = "UZZ" Then
            While Len(CStr(reqslno1)) < 6  ' CHANGED TO 6 FROM 7 AS HEMOTOLOGY
            ' NEEDED NUMBERING LIKE
            ' F05999999 FOR BF
            ' P05999999 FOR HL
            ' M05999999 FOR BM

                reqslno1 = "0" & reqslno1
            Wend
        'ElseIf Trim(dept) = "CCZ" Then
        ElseIf Trim(dept) = "CCH" Then
            '***** Commented and added by Shraddha 03/02/10 *****
            If Len(CStr(reqslno1)) = 1 Then
                zeros = "000" & CStr(reqslno1)
            ElseIf Len(CStr(reqslno1)) = 2 Then
                zeros = "00" & CStr(reqslno1)
            ElseIf Len(CStr(reqslno1)) = 3 Then
                zeros = "0" & CStr(reqslno1)
            Else
                zeros = reqslno1
            End If
            '****************************************************
        Else
            While Len(CStr(reqslno1)) < 6
                reqslno1 = "0" & reqslno1
            Wend
        End If
        DAUTOPATHNO = reqslno1
'******************************************************************
        'If Trim(MODALITY) = "SP" Then
        If Trim(DEPT1) = "UZZ" Then
            ' CHANGED TO 6 FROM 7 AS HEMOTOLOGY
            ' NEEDED NUMBERING LIKE
            ' F05999999 FOR BF
            ' P05999999 FOR HL
            ' M05999999 FOR BM
            SPAUTONO = Trim(PATH) & CStr(Trim(Right(DATE11, 2))) & Trim(DAUTOPATHNO)
        Else
            SPAUTONO = DAUTOPATHNO & "/" & PATH
        End If
        'End If

     txt_pathno.Text = SPAUTONO
Exit Function
'MSG:
'MsgBox Err.Description
End Function
Function IPNO() As String
Dim PATH As String
Dim MODALITY As String
MODALITY = "IP"
'On Error GoTo MSG:
DEPT1 = "UZZ"
Dim SELDAUTOPATH_QR As rdoQuery
Dim SELDAUTOPATH_RS As rdoResultset
Dim SELDAUTOPATHFLAG As Boolean

' ***** AUTOGENERATION OF THE PATHNO.
str = "SELECT * FROM DISLIB.DAUTOPATH WHERE DEPT_CODE=? AND MODALITYCD=? AND PATHYEAR1=?"
Set SELDAUTOPATH_QR = rdb.CreateQuery("", str)

Dim qr_seldmulproce_U753 As rdoQuery
Dim RS_seldmulproce_U753 As rdoResultset
str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=? AND (PRO_CODE='U753' OR PRO_CODE='U754')"
Set qr_seldmulproce_U753 = rdb.CreateQuery("", str)


Dim INSDAUTOPATH_QR As rdoQuery
str = "INSERT INTO DISLIB.DAUTOPATH (DEPT_CODE,MODALITYCD,PATHYEAR,PATHYEAR1,PATHNO) VALUES (?,?,?,?,?)"
Set INSDAUTOPATH_QR = rdb.CreateQuery("", str)

Dim UPDDAUTOPATH_QR As rdoQuery
str = "UPDATE DISLIB.DAUTOPATH SET PATHNO=? WHERE DEPT_CODE=? AND MODALITYCD=? AND PATHYEAR1=?"

Set UPDDAUTOPATH_QR = rdb.CreateQuery("", str)


Dim DPATHYEAR_QR As rdoQuery
Dim DPATHYEAR_RS As rdoResultset
Dim DPATHYEARFLAG As Boolean


str = "SELECT * FROM DISLIB.DPATHYEAR WHERE AYEAR=?"
Set DPATHYEAR_QR = rdb.CreateQuery("", str)


i = 1

           SELDAUTOPATH_QR(2) = Trim(Right(DATE11, 2))
           qr_seldmulproce_U753(0) = Replace(Trim(txt_reqno.Text), "/", "")
            Set RS_seldmulproce_U753 = qr_seldmulproce_U753.OpenResultset(2, 3)
If RS_seldmulproce_U753.EOF = True Then
Exit Function
End If

        If SELDAUTOPATHFLAG = False Then
            SELDAUTOPATH_QR(0) = Trim(DEPT1)
            If RS_seldmulproce_U753.EOF = False Then
            SELDAUTOPATH_QR(1) = "SM"
            Else
            SELDAUTOPATH_QR(1) = Trim(MODALITY)
            End If
            SELDAUTOPATH_QR(2) = Trim(Right(DATE11, 2))
            Set SELDAUTOPATH_RS = SELDAUTOPATH_QR.OpenResultset(2, 3)
            SELDAUTOPATHFLAG = True
        Else
            SELDAUTOPATH_QR(0) = Trim(DEPT1)
            'SELDAUTOPATH_QR(1) = Trim(MODALITY)
            If RS_seldmulproce_U753.EOF = False Then
            SELDAUTOPATH_QR(1) = "SM"
            Else
            SELDAUTOPATH_QR(1) = Trim(MODALITY)
            End If
            SELDAUTOPATH_QR(2) = Trim(Right(DATE11, 2))
            SELDAUTOPATH_RS.MoreResults
            SELDAUTOPATH_RS.Requery
        End If
        If SELDAUTOPATH_RS.RowCount > 0 Then
            AUTOPATHNO = SELDAUTOPATH_RS("PATHNO") + 1
            UPDDAUTOPATH_QR(0) = AUTOPATHNO
            UPDDAUTOPATH_QR(1) = Trim(DEPT1)
            'UPDDAUTOPATH_QR(2) = Trim(MODALITY)
            If RS_seldmulproce_U753.EOF = False Then
            UPDDAUTOPATH_QR(2) = "SM"
            Else
            UPDDAUTOPATH_QR(2) = Trim(MODALITY)
            End If
            UPDDAUTOPATH_QR(3) = Trim(Right(DATE11, 2))
            UPDDAUTOPATH_QR.Execute
        Else
            AUTOPATHNO = 1
            
            INSDAUTOPATH_QR(0) = Trim(DEPT1)
            'INSDAUTOPATH_QR(1) = Trim(MODALITY)
            If RS_seldmulproce_U753.EOF = False Then
            INSDAUTOPATH_QR(1) = "SM"
            Else
            INSDAUTOPATH_QR(1) = Trim(MODALITY)
            End If
            INSDAUTOPATH_QR(2) = Trim(Left(msk_caseno.FormattedText, 2))
            INSDAUTOPATH_QR(3) = Trim(Right(DATE11, 2))
            INSDAUTOPATH_QR(4) = 1
            INSDAUTOPATH_QR.Execute
        End If
        If Trim(DEPT1) = "UZZ" Then
                PATH = Trim(SELDAUTOPATH_RS("PATHYEAR"))
        Else
            If DPATHYEARFLAG = False Then
                DPATHYEAR_QR(0) = "A" & Right(DATE11, 4)
                Set DPATHYEAR_RS = DPATHYEAR_QR.OpenResultset(1, 3)
                DPATHYEARFLAG = True
            Else
                DPATHYEAR_QR(0) = "A" & Right(DATE11, 4)
                DPATHYEAR_RS.MoreResults
                DPATHYEAR_RS.Requery
            End If
            If DPATHYEAR_RS.RowCount > 0 Then
                PATH = DPATHYEAR_RS("LASTTWO")
            End If
        End If
    '***********padding the reqslno with leading zeros"****************
        Dim reqslno1
        reqslno1 = AUTOPATHNO
        If Trim(DEPT1) = "UZZ" Then
            While Len(CStr(reqslno1)) < 6  ' CHANGED TO 6 FROM 7 AS HEMOTOLOGY
            ' NEEDED NUMBERING LIKE
            ' F05999999 FOR BF
            ' P05999999 FOR HL
            ' M05999999 FOR BM

                reqslno1 = "0" & reqslno1
            Wend
        'ElseIf Trim(dept) = "CCZ" Then
        ElseIf Trim(dept) = "CCH" Then
            '***** Commented and added by Shraddha 03/02/10 *****
            If Len(CStr(reqslno1)) = 1 Then
                zeros = "000" & CStr(reqslno1)
            ElseIf Len(CStr(reqslno1)) = 2 Then
                zeros = "00" & CStr(reqslno1)
            ElseIf Len(CStr(reqslno1)) = 3 Then
                zeros = "0" & CStr(reqslno1)
            Else
                zeros = reqslno1
            End If
            '****************************************************
        Else
            While Len(CStr(reqslno1)) < 6
                reqslno1 = "0" & reqslno1
            Wend
        End If
        DAUTOPATHNO = reqslno1
'******************************************************************
        'If Trim(MODALITY) = "SP" Then
        
Dim qr_seldmulproce_U753A As rdoQuery
Dim RS_seldmulproce_U753A As rdoResultset
str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=? AND (PRO_CODE='U753' OR PRO_CODE='U754')"
Set qr_seldmulproce_U753A = rdb.CreateQuery("", str)
qr_seldmulproce_U753A(0) = Replace(Trim(txt_reqno.Text), "/", "")
Set RS_seldmulproce_U753A = qr_seldmulproce_U753A.OpenResultset(2, 3)
 If RS_seldmulproce_U753A.EOF = False Then
        If Trim(DEPT1) = "UZZ" Then
            ' CHANGED TO 6 FROM 7 AS HEMOTOLOGY
            ' NEEDED NUMBERING LIKE
            ' F05999999 FOR BF
            ' P05999999 FOR HL
            ' M05999999 FOR BM
            IPNO = Trim(DAUTOPATHNO)
        Else
            IPNO = DAUTOPATHNO
        End If
        'End If
        txt_ipno.Enabled = True
        txt_ipno.Text = ""
        
     LO = "IP" & Format(DATE11, "YY") & DAUTOPATHNO
     txt_ipno.Text = LO
     End If
Exit Function
MSG:
'MsgBox Err.Description
End Function
'***** This SPAUTONO function was Added by Shraddha on 12/10/09
Public Sub receiveNo(ByVal REQNO As String)
'Dim PATH As String
''On Error GoTo MSG:
'i = 1
'Dim modl As String
'DEPT1 = Left(txt_reqno.Text, 3)
'modl = Mid(txt_reqno.Text, 4, 2)
'
'    If Trim(txt_reqno.Text) <> "" Then
'
'            SELDAUTOPATH_QR(0) = Trim(dept)
'            'SELDAUTOPATH_QR(1) = Trim(MODALITY)
'            SELDAUTOPATH_QR(1) = modl ' Trim(SELDEPT_RS("MODALITYCD"))
'            SELDAUTOPATH_QR(2) = Trim(Right(DATE11, 2))
'            Set SELDAUTOPATH_RS = SELDAUTOPATH_QR.OpenResultset(2, 3)
'
'
'
'        If SELDAUTOPATH_RS.RowCount > 0 Then
'            AUTOPATHNO = SELDAUTOPATH_RS("PATHNO") + 1
'            UPDDAUTOPATH_QR(0) = AUTOPATHNO
'            UPDDAUTOPATH_QR(1) = Trim(dept)
'            UPDDAUTOPATH_QR(2) = Trim(SELDEPT_RS("MODALITYCD"))
'            UPDDAUTOPATH_QR(3) = Trim(Right(DATE11, 2))
'            UPDDAUTOPATH_QR.Execute
'        Else
'            AUTOPATHNO = 1
'            INSDAUTOPATH_QR(0) = Trim(dept)
'            INSDAUTOPATH_QR(1) = Trim(MODALITYB)
'            INSDAUTOPATH_QR(2) = Trim(Left(MSFlexGrid1.TextMatrix(i, 2), 2))
'            INSDAUTOPATH_QR(3) = Trim(Right(DATE11, 2))
'            INSDAUTOPATH_QR(4) = 1
'            INSDAUTOPATH_QR.Execute
'        End If
'        If Trim(dept) = "UZZ" Then
'                PATH = Trim(SELDAUTOPATH_RS("PATHYEAR"))
'        Else
'
'            DPATHYEAR_QR(0) = "A" & Right(DATE11, 4)
'            Set DPATHYEAR_RS = DPATHYEAR_QR.OpenResultset(1, 3)
'
'            If DPATHYEAR_RS.RowCount > 0 Then
'                PATH = DPATHYEAR_RS("LASTTWO")
'            End If
'        End If
'    '***********padding the reqslno with leading zeros"****************
'        Dim reqslno1
'        reqslno1 = AUTOPATHNO
'        If Trim(dept) = "UZZ" Then
'            While Len(CStr(reqslno1)) < 6  ' CHANGED TO 6 FROM 7 AS HEMOTOLOGY
'            ' NEEDED NUMBERING LIKE
'            ' F05999999 FOR BF
'            ' P05999999 FOR HL
'            ' M05999999 FOR BM
'
'                reqslno1 = "0" & reqslno1
'            Wend
'        Else
'            While Len(CStr(reqslno1)) < 6
'                reqslno1 = "0" & reqslno1
'            Wend
'        End If
'        DAUTOPATHNO = reqslno1
''******************************************************************
'        'If Trim(MODALITY) = "SP" Then
'        If Trim(dept) = "UZZ" Then
'            ' CHANGED TO 6 FROM 7 AS HEMOTOLOGY
'            ' NEEDED NUMBERING LIKE
'            ' F05999999 FOR BF
'            ' P05999999 FOR HL
'            ' M05999999 FOR BM
'            'MSFlexGrid1.TextMatrix(I, 7) = Trim(PATH) & CStr(Trim(Right(DATE11, 2))) & Trim(DAUTOPATHNO)
'            'Commented and added by Shraddha on 09/12/09
'            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = Trim(PATH) & CStr(Trim(Right(DATE11, 2))) & Trim(DAUTOPATHNO)
'        Else
'            'MSFlexGrid1.TextMatrix(I, 7) = DAUTOPATHNO & "/" & PATH
'            'Commented and added by Shraddha on 09/12/09
'            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = DAUTOPATHNO & "/" & PATH
'        End If
'
'
'        upddrequisitn(0) = UCase(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7)))
'        upddrequisitn(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
'        upddrequisitn.Execute
'
'        If Trim(Mid(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1), 4, 2)) = "HL" Then
'            upd_TEMMPCOLL_QR(0) = UCase(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7)))
'            upd_TEMMPCOLL_QR(1) = ipadd
'            upd_TEMMPCOLL_QR(2) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
'            upd_TEMMPCOLL_QR.Execute
'
'
'        End If
'        'End If
'    End If
'Next i
'Exit Sub
'MSG:
'MsgBox Err.Description
End Sub


Private Sub txt_reqno_LostFocus()
Dim prpathno As String
Dim entr, CytoFlag As String
'On Error GoTo MSG:
'' setting of tabs based on dhemcyto1 table
'      SSTab1.TabEnabled(0) = False
'      SSTab1.TabEnabled(1) = False
'      SSTab1.TabEnabled(2) = False
'      SSTab1.TabEnabled(3) = False
'      SSTab1.TabEnabled(4) = False

'************ Shraddha Pandit 30/03/09 ****************
'******************************************************

If Trim(msk_caseno.Text) = "" Then
'MsgBox "Case no is Blank", vbInformation
Exit Sub

Else

 '' seting of tabs based on dhemcyto1 table over
If Opt_patient.Value = True Or Opt_patient.Value = False Then
        If NEWFLAG = True Then
            NEWFLAG = False
            If selta101flag1 = False Then
                SELTA101_QR(0) = msk_caseno.FormattedText
                Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
                selta101flag1 = True
            Else
                SELTA101_QR(0) = msk_caseno.FormattedText
                SELTA101_RS.MoreResults
                SELTA101_RS.Requery
            End If
            If SELTA101_RS.RowCount = 0 Then
                Exit Sub
            End If
        datetime
            txt_sex.Text = SELTA101_RS("sex")
            txt_name.Text = Trim(SELTA101_RS("first_name")) & " " & Trim(SELTA101_RS("mid_name")) & " " & Trim(SELTA101_RS("last_name"))
            If Len(Trim(SELTA101_RS("BIRTH_DATE"))) <> 0 Then
                    'PATYPE = Trim(SELTA101_RS("PAT_TYPE"))
                   YEAR1 = DateDiff("m", SELTA101_RS("BIRTH_DATE"), DATE11)

                   Dim n As Integer
                   n = YEAR1 / 12
                   txt_age.Text = n
               End If

               If Trim(txt_age.Text) = 0 Then
                  Dim age As String
                  age = DateDiff("m", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                  age = "." & (age Mod 12)
                     If age = ".0" Then
                        age = DateDiff("d", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                        age = ".0" & (age) ' Mod 30)
                     End If
                  txt_age.Text = Trim(age)
               End If
            'txt_age = DateDiff("YYYY", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
          '  txt_pridoc.Text = SELTA101_RS("ud_name")
            txt_category.Text = SELTA101_RS("pat_catg")


            If SELMA104FLAG1 = False Then
                selma104_qr(0) = msk_caseno.FormattedText
                Set selma104_rs = selma104_qr.OpenResultset(1, 3)
                SELMA104FLAG1 = True
            Else
                selma104_qr(0) = msk_caseno.FormattedText
                selma104_rs.MoreResults
                selma104_rs.Requery
            End If

            If selma104_rs.RowCount > 0 Then
                txt_ward.Visible = True
                txt_bedno.Visible = True
                lbl_ward.Visible = True
                lbl_bed.Visible = True
                txt_ward.Text = selma104_rs("WARD_NAME")
                txt_bedno.Text = selma104_rs("bed_no")
                Text1.Text = "IP"
            Else
                Text1.Text = "OP"
                txt_ward.Visible = False
                txt_bedno.Visible = False
                lbl_ward.Visible = False
                lbl_bed.Visible = False
            End If

''
End If
End If

sel_rtime_qr(0) = Trim(Replace(txt_reqno.Text, "/", ""))
    Set sel_rtime_rs = sel_rtime_qr.OpenResultset(1, 3)

    If Not sel_rtime_rs.EOF Then
        TXT_COLLDTTM.Text = Trim(sel_rtime_rs("SAMCOLDATE")) & " / " & Trim(sel_rtime_rs("SAMCOLTIME"))
        TXT_RCVDTTM.Text = DATE11 & " / " & time1 '  Trim(sel_rtime_rs("SAMRECDATE")) & " / " & Trim(sel_rtime_rs("SAMRECTIME"))
    End If
'''receiving Number

'receiveNo (Trim(Replace(txt_reqno.Text, "/", "")))

'--------21-11-2013 increament path no on save button----------
'SPAUTONO
'IPNO

End If
End Sub

Private Sub TXT_SLNO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TXT_SLNO_LostFocus
Else
If KeyAscii > 47 And KeyAscii < 58 Then
Else
KeyAscii = 0
End If

End If
End Sub

Private Sub TXT_SLNO_LostFocus()
If TXT_SLNO.Text = "" Then
Exit Sub
End If
If TXT_SLNO.Text > txt_printQty.Text Then
MsgBox "Entered No. is greater than Quantity", vbCritical
TXT_SLNO.SetFocus
Exit Sub
End If
slno1 = TXT_SLNO.Text

MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 0) = TXT_SLNO.Text

End Sub
