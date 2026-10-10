VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frm_MLreqn 
   BackColor       =   &H80000003&
   Caption         =   "MEDICAL ONCOLOGY MOLECULAR LABORATORY"
   ClientHeight    =   10515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13440
   LinkTopic       =   "Form2"
   ScaleHeight     =   10515
   ScaleWidth      =   13440
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chk_Mutation 
      BackColor       =   &H80000003&
      Caption         =   "Mutation Analysis by ASO PCR "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   19
      Top             =   9000
      Width           =   6135
   End
   Begin VB.CheckBox chk_Gene 
      BackColor       =   &H80000003&
      Caption         =   "Gene Rearrangement"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1920
      TabIndex        =   18
      Top             =   8520
      Width           =   5775
   End
   Begin VB.CheckBox chk_Nested 
      BackColor       =   &H80000003&
      Caption         =   "RT-PCR Nested"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   17
      Top             =   7920
      Width           =   4095
   End
   Begin VB.CheckBox chk_HotStart 
      BackColor       =   &H80000003&
      Caption         =   "RT-PCR Hot Start"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1920
      TabIndex        =   16
      Top             =   7440
      Width           =   4575
   End
   Begin VB.CheckBox chk_PCRnestedTCR 
      BackColor       =   &H80000003&
      Caption         =   "RT-PCR Nested, TCR Gene Rearrangement"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   15
      Top             =   6840
      Width           =   4455
   End
   Begin VB.CheckBox chk_PCRnestedIGH 
      BackColor       =   &H80000003&
      Caption         =   "RT-PCR Nested, IGH Chain Gene Rearrangement "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   14
      Top             =   6240
      Width           =   4455
   End
   Begin VB.CheckBox chk_PCRPML 
      BackColor       =   &H80000003&
      Caption         =   "RQ-PCR PML-RARA"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   13
      Top             =   5640
      Width           =   4215
   End
   Begin VB.CheckBox chk_PCRMultiplexALP 
      BackColor       =   &H80000003&
      Caption         =   "RT-PCR Multiplex, Acute Leukaemia Panel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   12
      Top             =   5040
      Width           =   4215
   End
   Begin VB.CheckBox chk_PCRBCR 
      BackColor       =   &H80000003&
      Caption         =   "RQ-PCR BCR-ABL (P210)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   11
      Top             =   4440
      Width           =   4215
   End
   Begin VB.CheckBox chk_PCRnested 
      BackColor       =   &H80000003&
      Caption         =   "RT-PCR Nested, BCR-ABL for Follow-Up"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   3840
      Width           =   4215
   End
   Begin VB.CheckBox chk_PCRMultiplex 
      BackColor       =   &H80000003&
      Caption         =   "RT-PCR Multiplex, BCR-ABL (P190, P210)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   3240
      Width           =   4095
   End
   Begin VB.ListBox CMB_NMETERIAL 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      ItemData        =   "frm_MLreqn.frx":0000
      Left            =   2280
      List            =   "frm_MLreqn.frx":000D
      Style           =   1  'Checkbox
      TabIndex        =   8
      Top             =   1920
      Width           =   3255
   End
   Begin VB.CommandButton cmd_back 
      Caption         =   "<< BACK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   9840
      Width           =   1695
   End
   Begin VB.TextBox txt_name 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4080
      TabIndex        =   1
      Top             =   1200
      Width           =   5175
   End
   Begin MSMask.MaskEdBox Msk_caseno 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      ForeColor       =   0
      PromptInclude   =   0   'False
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
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
   Begin VB.Shape Shape2 
      Height          =   1935
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   9255
   End
   Begin VB.Shape Shape1 
      Height          =   6495
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   9255
   End
   Begin VB.Label lbl_specimen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nature of Metarial"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   600
      TabIndex        =   7
      Top             =   1920
      Width           =   1515
   End
   Begin VB.Label lbl_caseno 
      BackStyle       =   0  'Transparent
      Caption         =   "Case No."
      BeginProperty Font 
         Name            =   "Times New Roman"
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
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lbl_name 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3360
      TabIndex        =   5
      Top             =   1200
      Width           =   465
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "MEDICAL ONCOLOGY MOLECULAR LABORATORY"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   480
      Width           =   6615
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "frm_MLreqn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String
Dim sel_sercode_qr As rdoQuery
Dim sel_sercode_rs As rdoResultset

Private Sub Check1_Click()

End Sub



Private Sub Check4_Click()

End Sub

Private Sub Form_Load()
'str = "SELECT * FROM PATOOLSLIB.MA101"
'Set sel_sercode_qr = rdb.CreateQuery("", str)
'
'Set sel_sercode_rs = sel_sercode_qr.OpenResultset(1, 3)
'    While Not sel_sercode_rs.EOF
'        cmb_dept.AddItem Trim(sel_sercode_rs("dept_name")) & "$" & Trim(sel_sercode_rs("dept_code"))
'        sel_sercode_rs.MoveNext
'    Wend
End Sub

Private Sub cmd_back_Click()

'+++++++++
'If lst_itemcode.SelCount <> 0 Then
'
'    For I = 0 To lst_itemcode.ListCount - 1
'        If lst_itemcode.Selected(I) = True Then
'
'            KL = InStr(1, Trim(lst_itemcode.List(I)), "-")
'            Qr_insert(3) = Mid(Trim(lst_itemcode.List(I)), 1, KL - 1)
'
'        End If
'    Next I
'
'End If
'+++++++++

meterial = ""
If CMB_NMETERIAL.SelCount <> 0 Then
    For i = 0 To CMB_NMETERIAL.ListCount - 1
        If CMB_NMETERIAL.Selected(i) = True Then
            'meterial = CMB_NMETERIAL.Text
            meterial = CMB_NMETERIAL.List(i) & "," & meterial
        End If
    Next i
End If

If CMB_NMETERIAL.SelCount = 0 Then
    MsgBox "Please select nature of material"
    Exit Sub
End If

pathology.CMB_TEXT.clear
ccount = 0



'If chk_calc.Value = 1 Then
'    If ccount > 11 Then
'        MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
'        Unload Me
'        pathology.show
'        Exit Sub
'    End If
'    pathology.CMB_TEXT.AddItem "F888"
'    ccount = ccount + 1
'End If


'chk_PCRMultiplex
'chk_PCRnested
'chk_PCRBCR
'chk_PCRMultiplexALP
'chk_PCRPML
'chk_PCRnestedIGH
'chk_PCRnestedTCR

If chk_PCRMultiplex.Value = 1 Then
    pathology.CMB_TEXT.AddItem "J604"
End If

If chk_PCRnested.Value = 1 Then
    pathology.CMB_TEXT.AddItem "J605"
End If

If chk_PCRBCR.Value = 1 Then
    pathology.CMB_TEXT.AddItem "J606"
End If

If chk_PCRMultiplexALP.Value = 1 Then
    pathology.CMB_TEXT.AddItem "J607"
End If

If chk_PCRPML.Value = 1 Then
    pathology.CMB_TEXT.AddItem "J608"
End If

If chk_PCRnestedIGH.Value = 1 Then
    pathology.CMB_TEXT.AddItem "J609"
End If

If chk_PCRnestedTCR.Value = 1 Then
    pathology.CMB_TEXT.AddItem "J610"
End If

'***** Shraddha 08/04/10 *****
If chk_HotStart.Value = 1 Then
    pathology.CMB_TEXT.AddItem "J611"
End If

If chk_Nested.Value = 1 Then
    pathology.CMB_TEXT.AddItem "J612"
End If

If chk_Gene.Value = 1 Then
    pathology.CMB_TEXT.AddItem "J613"
End If

If chk_Mutation.Value = 1 Then
    pathology.CMB_TEXT.AddItem "J614"
End If
'*****************************

If pathology.CMB_TEXT.ListCount = 0 Then
    MsgBox ("Please atleast 1 procedure code"), vbInformation
    Exit Sub
End If


Unload Me
pathology.show
End Sub
