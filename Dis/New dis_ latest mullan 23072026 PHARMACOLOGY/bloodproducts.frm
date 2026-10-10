VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form bloodProducts 
   Caption         =   "Blood Products"
   ClientHeight    =   8700
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14115
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
   ScaleHeight     =   8700
   ScaleWidth      =   14115
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_name 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5160
      TabIndex        =   47
      Top             =   1200
      Width           =   5655
   End
   Begin VB.CommandButton cmd_back 
      Caption         =   "<< &Back"
      Height          =   495
      Left            =   9120
      TabIndex        =   30
      Top             =   8040
      Width           =   1455
   End
   Begin VB.TextBox TXT_NOUREPLACE 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   29
      Top             =   7200
      Width           =   2895
   End
   Begin VB.TextBox txt_phname 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   28
      Top             =   7200
      Width           =   2655
   End
   Begin VB.Frame Frame3 
      Caption         =   "Special Requirements"
      Height          =   975
      Left            =   720
      TabIndex        =   41
      Top             =   6000
      Width           =   10095
      Begin VB.TextBox TXT_OTHERS 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5040
         TabIndex        =   27
         Top             =   360
         Width           =   4455
      End
      Begin VB.CheckBox CHK_LEUCODEPLETION 
         Caption         =   "Leucodepletion"
         Height          =   255
         Left            =   2040
         TabIndex        =   26
         Top             =   360
         Width           =   1935
      End
      Begin VB.CheckBox CHK_IRRADIATION 
         Caption         =   "Irradiation"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lable9 
         AutoSize        =   -1  'True
         Caption         =   "Remarks"
         Height          =   195
         Left            =   4200
         TabIndex        =   42
         Top             =   360
         Width           =   750
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Products Required"
      Height          =   1935
      Left            =   720
      TabIndex        =   40
      Top             =   3840
      Width           =   10095
      Begin VB.ComboBox cmb_pcunit 
         Height          =   315
         ItemData        =   "bloodproducts.frx":0000
         Left            =   2880
         List            =   "bloodproducts.frx":000D
         TabIndex        =   58
         Top             =   1440
         Width           =   1095
      End
      Begin VB.ComboBox cmb_wbunit 
         Height          =   315
         ItemData        =   "bloodproducts.frx":0029
         Left            =   840
         List            =   "bloodproducts.frx":0036
         TabIndex        =   57
         Top             =   1440
         Width           =   1095
      End
      Begin VB.ComboBox cmb_aliquot 
         Height          =   315
         ItemData        =   "bloodproducts.frx":0052
         Left            =   4320
         List            =   "bloodproducts.frx":0054
         TabIndex        =   56
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txt_sdp 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         MaxLength       =   2
         TabIndex        =   16
         Top             =   1440
         Width           =   495
      End
      Begin VB.CheckBox chk_sdp 
         Caption         =   "SDP"
         Height          =   255
         Left            =   5160
         TabIndex        =   15
         ToolTipText     =   "Single Donor Platelets"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txt_alb 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9240
         MaxLength       =   2
         TabIndex        =   24
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txt_cryo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8160
         MaxLength       =   2
         TabIndex        =   22
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txt_fviiid 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7080
         MaxLength       =   2
         TabIndex        =   20
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txt_ffp 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6120
         MaxLength       =   2
         TabIndex        =   18
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txt_rdp 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   14
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txt_pc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   12
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txt_wb 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         MaxLength       =   2
         TabIndex        =   10
         Top             =   1440
         Width           =   495
      End
      Begin VB.CheckBox chk_alb 
         Caption         =   "ALB"
         Height          =   255
         Left            =   9240
         TabIndex        =   23
         ToolTipText     =   "Albumin"
         Top             =   840
         Width           =   735
      End
      Begin VB.CheckBox chk_cryo 
         Caption         =   "CRYO"
         Height          =   255
         Left            =   8160
         TabIndex        =   21
         ToolTipText     =   "Cryoprecipitate"
         Top             =   840
         Width           =   855
      End
      Begin VB.CheckBox chk_fviiid 
         Caption         =   "FVIIID"
         Height          =   195
         Left            =   7080
         TabIndex        =   19
         ToolTipText     =   "Factor VIII Deficient Plasma"
         Top             =   840
         Width           =   975
      End
      Begin VB.CheckBox chk_ffp 
         Caption         =   "FFP"
         Height          =   255
         Left            =   6120
         TabIndex        =   17
         ToolTipText     =   "Fresh Frozen Plasma"
         Top             =   840
         Width           =   735
      End
      Begin VB.CheckBox chk_rdp 
         Caption         =   "RDP"
         Height          =   255
         Left            =   4200
         TabIndex        =   13
         ToolTipText     =   "Random Donor Platelets"
         Top             =   840
         Width           =   735
      End
      Begin VB.CheckBox chk_pc 
         Caption         =   "PC"
         Height          =   255
         Left            =   2280
         TabIndex        =   11
         ToolTipText     =   "Packed Cells"
         Top             =   840
         Width           =   615
      End
      Begin VB.CheckBox chk_wb 
         Caption         =   "WB"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         ToolTipText     =   "Whole Blood"
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "Is an aliquot of blood required for your patient?"
         Height          =   375
         Left            =   240
         TabIndex        =   55
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   720
      TabIndex        =   32
      Top             =   1680
      Width           =   10095
      Begin VB.TextBox TXT_PLTC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8400
         MaxLength       =   10
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox cmb_time 
         Height          =   315
         Left            =   6120
         TabIndex        =   6
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox TXT_DTMNO 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   46
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox CMB_RT 
         Height          =   315
         ItemData        =   "bloodproducts.frx":0056
         Left            =   2400
         List            =   "bloodproducts.frx":0058
         TabIndex        =   2
         Top             =   720
         Width           =   4095
      End
      Begin VB.ComboBox cmb_tr 
         Height          =   315
         Left            =   8400
         TabIndex        =   8
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox TXT_TM 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5040
         TabIndex        =   5
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox TXT_DT 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   4
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox TXT_HB 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8400
         MaxLength       =   5
         TabIndex        =   7
         Top             =   1200
         Width           =   855
      End
      Begin VB.ComboBox cmb_bg1 
         Height          =   315
         Left            =   4920
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cmb_pr 
         Height          =   315
         ItemData        =   "bloodproducts.frx":005A
         Left            =   8400
         List            =   "bloodproducts.frx":005C
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label13 
         Caption         =   "Plt. Count (10E9/L)"
         Height          =   255
         Left            =   6720
         TabIndex        =   52
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "DTM No."
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Type of Requisition"
         Height          =   195
         Left            =   6720
         TabIndex        =   39
         Top             =   1680
         Width           =   1665
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Time"
         Height          =   195
         Left            =   4560
         TabIndex        =   38
         Top             =   1440
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Required on"
         Height          =   195
         Left            =   1200
         TabIndex        =   37
         Top             =   1440
         Width           =   1050
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Hb g/dl"
         Height          =   195
         Left            =   7680
         TabIndex        =   36
         Top             =   1200
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Reason for Transfusion"
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   840
         Width           =   1995
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Blood Group"
         Height          =   195
         Left            =   3600
         TabIndex        =   34
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label2 
         Caption         =   "Previous Recipent"
         Height          =   255
         Left            =   6720
         TabIndex        =   33
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSMask.MaskEdBox msk_caseno 
      Height          =   255
      Left            =   2280
      TabIndex        =   48
      Top             =   1200
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
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "HOMI BHABHA CANCER HOSPITAL AND RESEARCH CENTRE,MULLANPUR"
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
      Height          =   495
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Width           =   13935
   End
   Begin VB.Label Label14 
      Caption         =   "*"
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
      Left            =   600
      TabIndex        =   53
      Top             =   7320
      Width           =   135
   End
   Begin VB.Label Label12 
      Caption         =   "Printout should be Signed by Doctor ordering  Blood and the Phlebotomist."
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   840
      TabIndex        =   51
      Top             =   7920
      Width           =   7095
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
      TabIndex        =   50
      Top             =   1200
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
      Left            =   4440
      TabIndex        =   49
      Top             =   1200
      Width           =   630
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "No.of Units replaced"
      Height          =   195
      Left            =   5880
      TabIndex        =   44
      Top             =   7320
      Width           =   1770
   End
   Begin VB.Label Label9 
      Caption         =   "Phlebotomist Name"
      Height          =   255
      Left            =   840
      TabIndex        =   43
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "REQUISITION FOR BLOOD PRODUCTS"
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
      TabIndex        =   31
      Top             =   600
      Width           =   5655
   End
End
Attribute VB_Name = "bloodProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CHK_ALB_Click()
If CHK_ALB.Value = 1 Then
    TXT_ALB.Visible = True
    TXT_ALB.SetFocus
Else
    TXT_ALB.Visible = False
End If
End Sub

Private Sub CHK_CRYO_Click()
If CHK_CRYO.Value = 1 Then
    TXT_CRYO.Visible = True
    TXT_CRYO.SetFocus
Else
    TXT_CRYO.Visible = False
End If
End Sub

Private Sub CHK_FFP_Click()
If CHK_FFP.Value = 1 Then
    TXT_FFP.Visible = True
    TXT_FFP.SetFocus
Else
    TXT_FFP.Visible = False
End If
End Sub

Private Sub CHK_FVIIID_Click()
If CHK_FVIIID.Value = 1 Then
    TXT_FVIIID.Visible = True
    TXT_FVIIID.SetFocus
Else
    TXT_FVIIID.Visible = False
End If
End Sub

Private Sub CHK_PC_Click()
If CHK_PC.Value = 1 Then
    TXT_PC.Visible = True
    '******** Added & commented by Shraddha 26/03/09
    If cmb_aliquot.Text = "Yes" Then
        TXT_PC.Enabled = False
        TXT_PC.Text = 1
    End If
    If cmb_aliquot.Text = "No" Then
        TXT_PC.Enabled = True
        TXT_PC.Text = ""
    End If
'    txt_pc.SetFocus
Else
    TXT_PC.Visible = False
End If

If cmb_aliquot.Text = "Yes" And TXT_PC.Visible Then
    cmb_pcunit.Visible = True
Else
    cmb_pcunit.Visible = False
End If
End Sub

Private Sub CHK_RDP_Click()
If CHK_RDP.Value = 1 Then
    TXT_RDP.Visible = True
    TXT_RDP.SetFocus
Else
    TXT_RDP.Visible = False
End If
End Sub

Private Sub CHK_SDP_Click()
If CHK_SDP.Value = 1 Then
    TXT_SDP.Visible = True
    TXT_SDP.SetFocus
Else
    TXT_SDP.Visible = False
End If
End Sub

Private Sub CHK_WB_Click()
If CHK_WB.Value = 1 Then
    TXT_WB.Visible = True
    '******** Added & commented by Shraddha 26/03/09
    If cmb_aliquot.Text = "Yes" Then
        TXT_WB.Enabled = False
        TXT_WB.Text = 1
    End If
    If cmb_aliquot.Text = "No" Then
        TXT_WB.Enabled = True
        TXT_WB.Text = ""
    End If
    
    'txt_wb.SetFocus
Else
    TXT_WB.Visible = False
End If

If cmb_aliquot.Text = "Yes" And TXT_WB.Visible Then
    cmb_wbunit.Visible = True
Else
    cmb_wbunit.Visible = False
End If

End Sub

Private Sub cmb_time_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmd_back_Click()
Dim TM
TM = Trim(TXT_DT.Text)
If Len(TM) <> 10 Then
    MsgBox "Pls Enter Suitable Date !!", vbInformation, "BP"
    TXT_DT.SetFocus
    Exit Sub
End If
If Len(cmb_time.Text) = 0 Then
    MsgBox "PLEASE SELECT AM OR PM TIME FORMAT !!", vbInformation, "BP"
    cmb_time.SetFocus
    Exit Sub
End If

'********************* Shraddha Pandit 02/03/09 *********************
'chk_rdp chk_sdp
'If TXT_PLTC.Text = "" Then
'change again on 31/03/09
If (CHK_RDP.Value = 1 Or CHK_SDP.Value = 1) And TXT_PLTC.Text = "" Then
    MsgBox "Plt. Count field can not be Empty !!", vbCritical, "Blood Products"
    TXT_PLTC.Enabled = True
    TXT_PLTC.SetFocus
    Exit Sub
End If
'If TXT_HB.Text = "" Then
'change again on 31/03/09
If (CHK_WB.Value = 1 Or CHK_PC.Value = 1) And TXT_HB.Text = "" Then
    MsgBox "Hb g/dl field can not be Empty !!", vbCritical, "Blood Products"
    TXT_HB.Enabled = True
    TXT_HB.SetFocus
    Exit Sub
End If
'********************************************************************

If txt_dtmno.Text = "" Then
    'Commented by Shraddha on 11/08/10 *****
    'MsgBox "DTM No. field can not be Empty !!", vbInformation, "Blood Products"
    txt_dtmno.Enabled = True
    txt_dtmno.SetFocus
    Exit Sub
End If
'MsgBox "Pls Check DTM No.carefully !!", vbCritical, "Blood Products"
'TXT_DTMNO.SelStart = 0
'TXT_DTMNO.SelLength = Len(TXT_DTMNO)
'TXT_DTMNO.SetFocus
If Len(Trim(txt_phname.Text)) = 0 Then
    MsgBox "Please Enter the Phlebotomist Name"
    txt_phname.SetFocus
    Exit Sub
End If
If (CHK_WB.Value = 1 And Len(TXT_WB.Text) = 0) Then
    MsgBox "Please Enter the WB value"
    TXT_WB.SetFocus
    Exit Sub
End If

If (CHK_PC.Value = 1 And Len(TXT_PC.Text) = 0) Then
    MsgBox "Please Enter the PC value"
    TXT_PC.SetFocus
    Exit Sub
End If

If (CHK_RDP.Value = 1 And Len(TXT_RDP.Text) = 0) Then
    MsgBox "Please Enter the RDP value"
    TXT_RDP.SetFocus
    Exit Sub
End If

If (CHK_SDP.Value = 1 And Len(TXT_SDP.Text) = 0) Then
    MsgBox "Please Enter the SDP value"
    TXT_SDP.SetFocus
    Exit Sub
End If

If (CHK_FFP.Value = 1 And Len(TXT_FFP.Text) = 0) Then
    MsgBox "Please Enter the FFP value"
    TXT_FFP.SetFocus
    Exit Sub
End If

If (CHK_FVIIID.Value = 1 And Len(TXT_FVIIID.Text) = 0) Then
    MsgBox "Please Enter the FVIIID value"
    TXT_FVIIID.SetFocus
    Exit Sub
End If

If (CHK_CRYO.Value = 1 And Len(TXT_CRYO.Text) = 0) Then
    MsgBox "Please Enter the CYRO value"
    TXT_CRYO.SetFocus
    Exit Sub
End If

If (CHK_ALB.Value = 1 And Len(TXT_ALB.Text) = 0) Then
    MsgBox "Please Enter the ALB value"
    TXT_ALB.SetFocus
    Exit Sub
End If

ccount = 0
pathology.CMB_TEXT.clear
 ind1 = ""
 ind2 = ""
 ind3 = ""
 ind4 = ""
 ind5 = ""
 ind6 = ""
 ind7 = ""
 ind8 = ""
 ind9 = ""
 
 bpwb = ""
 bppc = ""
 bprdp = ""
 bpsdp = ""
 bpffd = ""
 bpfviiid = ""
 bpcryo = ""
 bpalb = ""
 
 BG = ""
 RFT = ""
 HB = ""
 REQUION = ""
 TM = ""
 TYPEOFREQ = ""
 SR1 = ""
 T_M = ""
 T_M1 = ""
 FLD = ""
 SR2 = ""
 OTHERS1 = ""
 PHOBOTMIST = ""
 NOOFUNREP = ""
 DTMNO = ""
 
If BPFLAG = False Then
    
    If CHK_WB.Value = 1 Or CHK_PC.Value = 1 Then
        pathology.CMB_TEXT.AddItem "H002"
        ccount = ccount + 1
        crossmatch = Val(TXT_WB.Text) + Val(TXT_PC.Text)
    End If
    
    If CHK_WB.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        pathology.CMB_TEXT.AddItem "H206"
        ind1 = TXT_WB.Text
        bpwb = "WHB"
    End If

    If CHK_PC.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        pathology.CMB_TEXT.AddItem "H207"
        ind2 = TXT_PC.Text
        bppc = "PC"
    End If
    ' ' sundeep for pimpol  inserting into Req of BP
        
        
         
    ' ' sundeep for pimpol  inserting into Req of BP
    If CHK_RDP.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        pathology.CMB_TEXT.AddItem "H210"
        ind3 = TXT_RDP.Text
        bprdp = "RDP"
    End If
    If CHK_SDP.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        pathology.CMB_TEXT.AddItem "H211"
        ind4 = Trim(TXT_SDP.Text)
        bpsdp = "SDP"
    End If
    If CHK_FFP.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        pathology.CMB_TEXT.AddItem "H209"
        ind5 = TXT_FFP.Text
        bpffp = "FFP"
    End If
    If CHK_FVIIID.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        pathology.CMB_TEXT.AddItem "?"
        ind6 = TXT_FVIIID.Text
        bpfviiid = "FVIIID"
    End If

    If CHK_CRYO.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        pathology.CMB_TEXT.AddItem "?"
        ind7 = TXT_CRYO.Text
        bpcryo = "CRYO"
    End If
    If CHK_ALB.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        pathology.CMB_TEXT.AddItem "H214"
        ind8 = TXT_ALB.Text
        bpalb = "ALB"
    End If
    If Len(Trim(TXT_PLTC.Text)) > 0 Then
        PLTC = Trim(TXT_PLTC.Text)
    End If
    If Len(Trim(txt_dtmno.Text)) > 0 Then
        DTMNO = Trim(txt_dtmno.Text)
    End If
    If Len(Trim(cmb_bg1.Text)) > 0 Then
        BG = cmb_bg1.Text
    End If
    If Len(Trim(CMB_RT.Text)) > 0 Then
        RFT = CMB_RT.Text
    End If
    If Len(Trim(TXT_HB.Text)) > 0 Then
        HB = Trim(TXT_HB.Text)
    End If
    
    
    '*********** Shraddha Pandit 06/03/09 ***********
    If Len(Trim(cmb_aliquot.Text)) > 0 Then
    aliquot = cmb_aliquot.Text
    End If
    
    If Len(Trim(cmb_wbunit.Text)) > 0 Then
    wbunit = cmb_wbunit.Text
    End If
    
    If Len(Trim(cmb_pcunit.Text)) > 0 Then
    pcunit = cmb_pcunit.Text
    End If
    '************************************************
    
    If Len(Trim(TXT_DT.Text)) > 0 Then
        REQUION = Trim(TXT_DT.Text)
    End If
    If Len(Trim(TXT_TM.Text)) > 0 Then
        T_M = Trim(TXT_TM.Text)
    End If
    If Len(Trim(cmb_time.Text)) > 0 Then
        T_M1 = Trim(cmb_time.Text)
    End If
    If Len(Trim(cmb_tr.Text)) > 0 Then
        TYPEOFREQ = Trim(cmb_tr.Text)
    End If
    If Trim(CHK_IRRADIATION.Value) = 1 Then
        SR1 = Trim(CHK_IRRADIATION.Caption)
    End If
    If Trim(CHK_LEUCODEPLETION.Value) = 1 Then
        SR2 = Trim(CHK_LEUCODEPLETION.Caption)
    End If
    If Len(Trim(TXT_OTHERS.Text)) > 0 Then
        OTHERS1 = Trim(TXT_OTHERS.Text)
    End If
    If Len(Trim(txt_phname.Text)) > 0 Then
        PHOBOTMIST = Trim(txt_phname.Text)
    End If
    If Len(Trim(TXT_NOUREPLACE.Text)) > 0 Then
        NOOFUNREP = Trim(TXT_NOUREPLACE.Text)
    End If
    If Len(Trim(cmb_time.Text)) > 0 Then
        TT = Trim(cmb_time.Text)
    End If
    If Len(Trim(CMB_PR.Text)) > 0 Then
        FLD = Trim(CMB_PR.Text)
    End If
End If
' ADD ON 28/06/2004
If CHK_WB.Value = 1 Or CHK_PC.Value = 1 Or CHK_RDP.Value = 1 Or CHK_SDP.Value = 1 Or CHK_FFP.Value = 1 Or CHK_FVIIID.Value = 1 Or CHK_CRYO.Value = 1 Or CHK_ALB.Value = 1 Then
If CHK_WB.Value = 1 Then
If TXT_WB.Text = "0" Or TXT_WB.Text = "" Then
    MsgBox "PLEASE ENTER PROPER NUMBER "
    Exit Sub '
End If
End If
If CHK_PC.Value = 1 Then
If TXT_PC.Text = "0" Or TXT_PC.Text = "" Then
    MsgBox "PLEASE ENTER PROPER NUMBER "
    Exit Sub
End If
End If
If CHK_RDP.Value = 1 Then
If TXT_RDP.Text = "" Or TXT_RDP.Text = "0" Then
    MsgBox "PLEASE ENTER PROPER NUMBER "
    Exit Sub
End If
End If
If CHK_SDP.Value = 1 Then
If TXT_SDP.Text = "" Or TXT_SDP.Text = "0" Then
    MsgBox "PLEASE ENTER PROPER NUMBER "
    Exit Sub
End If
End If
If CHK_FFP.Value = 1 Then
If TXT_FFP.Text = "" Or TXT_FFP.Text = "0" Then
    MsgBox "PLEASE ENTER PROPER NUMBER "
    Exit Sub
End If
End If
If CHK_FVIIID.Value = 1 Then
If TXT_FVIIID.Text = "" Or TXT_FVIIID.Text = "0" Then
    MsgBox "PLEASE ENTER PROPER NUMBER "
    Exit Sub
End If
End If
If CHK_CRYO.Value = 1 Then
If TXT_CRYO.Text = "" Or TXT_CRYO.Text = "0" Then
    MsgBox "PLEASE ENTER PROPER NUMBER "
    Exit Sub
End If
End If
If CHK_ALB.Value = 1 Then
If TXT_ALB.Text = "" Or TXT_ALB.Text = "0" Then
    MsgBox "PLEASE ENTER PROPER NUMBER "
    Exit Sub
End If
End If
End If

If CHK_WB.Value = 0 And CHK_PC.Value = 0 And CHK_SDP.Value = 0 And CHK_FFP.Value = 0 And CHK_FVIIID.Value = 0 And CHK_RDP.Value = 0 And CHK_ALB.Value = 0 And CHK_CRYO.Value = 0 Then
    MsgBox "Please Select atleast one value"
    Exit Sub
End If
BPFLAG = False
Unload Me
pathology.show
End Sub

Private Sub Form_Load()
cmb_tr.clear
cmb_tr.AddItem "Rush"
cmb_tr.AddItem "Urgent"
cmb_tr.AddItem "Planned"
cmb_tr.Text = "Planned"

cmb_bg1.clear
cmb_bg1.AddItem "A Pos"
cmb_bg1.AddItem "B Pos"
cmb_bg1.AddItem "O Pos"
cmb_bg1.AddItem "AB Pos"
cmb_bg1.AddItem "A Neg"
cmb_bg1.AddItem "B Neg"
cmb_bg1.AddItem "O Neg"
cmb_bg1.AddItem "AB Neg"
cmb_bg1.AddItem "Bombay Pos"
cmb_bg1.AddItem "Bombay Neg"

CMB_PR.clear
CMB_PR.AddItem "Yes"
CMB_PR.AddItem "No"
CMB_PR.Text = "No"

'*********** Shraddha Pandit 06/03/09 ***************
cmb_aliquot.clear
cmb_aliquot.AddItem "Yes"
cmb_aliquot.AddItem "No"
cmb_aliquot.Text = "No"

cmb_wbunit.Visible = False
cmb_pcunit.Visible = False
'****************************************************

cmb_time.AddItem "AM"
cmb_time.AddItem "PM"
TXT_WB.Visible = False
TXT_PC.Visible = False
TXT_RDP.Visible = False
TXT_SDP.Visible = False
TXT_FFP.Visible = False
TXT_FVIIID.Visible = False
TXT_CRYO.Visible = False
TXT_ALB.Visible = False

CMB_RT.clear
CMB_RT.AddItem "Anemia"
CMB_RT.AddItem "Anticipated bld. loss during Sx"
CMB_RT.AddItem "Intraoperative bleeding"
CMB_RT.AddItem "Deranged coagulation profile"
CMB_RT.AddItem "Thrombo cytopenia"
CMB_RT.AddItem "ANY OTHER:" 'added on 31/03/09 by Shraddha
cmb_bg1.Enabled = False
'txt_dtmno.SetFocus
End Sub


Private Sub txt_alb_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
End If
End Sub

Private Sub txt_cryo_Change()
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
End If
End Sub

Private Sub TXT_DTMNO_LostFocus()
If txt_dtmno.Text = "" Then
    MsgBox "DTM No. field can not be Empty !!", vbInformation, "Blood Products"
    txt_dtmno.SetFocus
    Exit Sub
End If
cmb_bg1.Enabled = False
End Sub

Private Sub txt_ffp_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
End If
End Sub

Private Sub txt_fviiid_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
End If
End Sub

Private Sub TXT_HB_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
End If
End Sub

Private Sub TXT_NOUREPLACE_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
End If
End Sub

Private Sub txt_pc_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
End If
End Sub


Private Sub txt_phname_LostFocus()
    If txt_dtmno.Enabled = True Then
        MsgBox "Pls Reconfirm DTM No.from Case File !!", vbInformation, "Blood Products"
         
        txt_dtmno.SelStart = 0
        txt_dtmno.SelLength = Len(txt_dtmno)
        txt_dtmno.SetFocus
    Else
    End If
End Sub

Private Sub TXT_PLTC_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
End If
End Sub

Private Sub txt_rdp_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
End If

End Sub

Private Sub txt_sdp_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
End If
End Sub

Private Sub TXT_TM_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 And KeyAscii <> 46 And KeyAscii <> 58 Then
      KeyAscii = 0
End If
End Sub

Private Sub txt_wb_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
End If
End Sub
