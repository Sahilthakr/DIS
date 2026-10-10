VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form HEMBF 
   Caption         =   "Requisition for Body Fluids"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12780
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
   ScaleHeight     =   7665
   ScaleWidth      =   12780
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_name 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4560
      TabIndex        =   16
      Top             =   1200
      Width           =   5655
   End
   Begin VB.CommandButton cmd_back 
      Caption         =   "<< &Back"
      Height          =   615
      Left            =   8880
      TabIndex        =   8
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   840
      TabIndex        =   1
      Top             =   4080
      Width           =   9375
      Begin VB.CheckBox CHK_SMC 
         Caption         =   "Surface Marker - Complete"
         Height          =   195
         Left            =   3600
         TabIndex        =   22
         Top             =   1080
         Width           =   2775
      End
      Begin VB.CheckBox CHK_BFC 
         Caption         =   "Body Fluids Cytospin"
         Height          =   255
         Left            =   6720
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CheckBox CHK_ANALYSIS 
         Caption         =   "Analysis"
         Height          =   195
         Left            =   480
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox CHK_CELLCOUNT 
         Caption         =   "Cell Count"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox chk_cytochemistry 
         Caption         =   "Cytochemistry"
         Height          =   255
         Left            =   4560
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox chk_sm 
         Caption         =   "Surface Markers - Single"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CheckBox chk_cytomorphology 
         Caption         =   "Cytomorphology"
         Height          =   255
         Left            =   2520
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nature of Material "
      Height          =   2295
      Left            =   840
      TabIndex        =   0
      Top             =   1680
      Width           =   9375
      Begin VB.CheckBox chk_pfa 
         Caption         =   "Pericardial Fluid Analysis"
         Height          =   255
         Left            =   3600
         TabIndex        =   21
         Top             =   960
         Width           =   2655
      End
      Begin VB.CheckBox chk_sa 
         Caption         =   "Semen Analysis"
         Height          =   255
         Left            =   7560
         TabIndex        =   20
         Top             =   1920
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox CHK_UPT 
         Caption         =   "Urine for Pregnancy test"
         Height          =   255
         Left            =   7800
         TabIndex        =   14
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox TXT_OTHERS 
         Appearance      =   0  'Flat
         Height          =   765
         Left            =   480
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   10
         Top             =   1440
         Width           =   5775
      End
      Begin VB.CheckBox chk_af 
         Caption         =   "Ascitic Fluid"
         Height          =   255
         Left            =   2040
         TabIndex        =   4
         Top             =   960
         Width           =   1455
      End
      Begin VB.CheckBox CHK_PF 
         Caption         =   "Pleural Fluid"
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   960
         Width           =   1455
      End
      Begin VB.CheckBox chk_csf 
         Caption         =   "CSF (Analysis : Cell Count + Cytomorphology)"
         Height          =   615
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   4335
      End
      Begin VB.Line Line1 
         Visible         =   0   'False
         X1              =   6480
         X2              =   6480
         Y1              =   120
         Y2              =   2280
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Others"
         Height          =   195
         Left            =   480
         TabIndex        =   11
         Top             =   1200
         Width           =   570
      End
   End
   Begin MSMask.MaskEdBox msk_caseno 
      Height          =   255
      Left            =   1800
      TabIndex        =   17
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   23
      Top             =   240
      Width           =   11655
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
      Left            =   840
      TabIndex        =   19
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
      Left            =   3840
      TabIndex        =   18
      Top             =   1200
      Width           =   630
   End
   Begin VB.Label lbl_cap 
      AutoSize        =   -1  'True
      Caption         =   "BODY FLUIDS - REQUISITION"
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
      Left            =   3720
      TabIndex        =   9
      Top             =   720
      Width           =   4110
   End
End
Attribute VB_Name = "HEMBF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chk_af_Click()
If chk_af.Value = 1 Then
CHK_ANALYSIS.Visible = False
    CHK_CELLCOUNT.Visible = False
    chk_cytomorphology.Visible = False

End If
End Sub

Private Sub CHK_ANALYSIS_Click()
If CHK_ANALYSIS.Value = 1 Then
    CHK_CELLCOUNT.Value = 1
    chk_cytomorphology.Value = 1
'    chk_cytochemistry.Value = 1
    CHK_BFC.Value = 1
'    chk_sm.Value = 1
Else
    CHK_CELLCOUNT.Value = 0
    chk_cytomorphology.Value = 0
'    chk_cytochemistry.Value = 0
    CHK_BFC.Value = 0
'    chk_sm.Value = 0
End If
End Sub

Private Sub chk_csf_Click()
If chk_csf.Value = 1 Then
    CHK_ANALYSIS.Visible = False
    CHK_CELLCOUNT.Visible = False
    chk_cytomorphology.Visible = False
Else
'    CHK_ANALYSIS.Visible = True
'    CHK_CELLCOUNT.Visible = True
'    chk_cytomorphology.Visible = True
End If
End Sub

Private Sub chk_cytochemistry_Click()
If chk_cytochemistry.Value = 1 Then
    chk_cytomorphology.Value = 1
Else
    chk_cytomorphology.Value = 0
End If
End Sub

Private Sub CHK_PF_Click()
If CHK_PF.Value = 1 Then
CHK_ANALYSIS.Visible = False
    CHK_CELLCOUNT.Visible = False
    chk_cytomorphology.Visible = False

End If
End Sub

Private Sub chk_pfa_Click()
If chk_pfa.Value = 1 Then
CHK_ANALYSIS.Visible = False
    CHK_CELLCOUNT.Visible = False
    chk_cytomorphology.Visible = False
End If
End Sub

'shweta 19-jan-2012
Private Sub chk_sm_Click()
If chk_sm.Value = 1 Then
CHK_SMC.Enabled = False
Else
CHK_SMC.Enabled = True
End If
End Sub

Private Sub CHK_SMC_Click()
If CHK_SMC.Value = 1 Then
chk_sm.Enabled = False
Else
chk_sm.Enabled = True
End If
End Sub
'''''''
Private Sub cmd_back_Click()
pathology.CMB_TEXT.clear
ccount = 0

ind1 = ""
ind2 = ""
ind3 = ""
ind4 = ""
ind5 = ""
'If BFFLAG = False Then
        '**** csf
        If chk_csf.Value = 1 Then
            If ccount > 11 Then
                MsgBox "You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "U718"
            meterial = meterial & "  " & Trim(chk_csf.Caption)
            ccount = ccount + 1
        End If
        '**** csf
        '**** pleural fluid
        If CHK_PF.Value = 1 Then
            If ccount > 11 Then
                MsgBox "You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "U726"
            meterial = meterial & "  " & Trim(CHK_PF.Caption)
            ccount = ccount + 1
        End If
        '**** pleural fluid
        '**** ascitic fluid
        If chk_af.Value = 1 Then
            If ccount > 11 Then
                MsgBox "You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "U725"
            meterial = meterial & "  " & Trim(chk_af.Caption)
            ccount = ccount + 1
        End If
        '**** ascitic  fluid
        '**** pericardical  fluid
        If chk_pfa.Value = 1 Then
            If ccount > 11 Then
                MsgBox "You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "U727"
            meterial = meterial & "  " & Trim(chk_pfa.Caption)
            ccount = ccount + 1
        End If
        '**** pericardial  fluid
        If txt_others.Text <> "" Then
                meterial = meterial & "  " & Trim(txt_others.Text)
        End If
         '**** upt  fluid
        If CHK_UPT.Value = 1 Then
            If ccount > 11 Then
                MsgBox "You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "U716"
            ccount = ccount + 1
        End If
        '**** upt  fluid
         '**** semen analysis
         If chk_sa.Value = 1 Then
            If ccount > 11 Then
                MsgBox "You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "U717"
            ccount = ccount + 1
        End If
        '**** semen analysis
        If chk_cytomorphology = 1 Then
            If ccount > 11 Then
                MsgBox "You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "U751"
            ccount = ccount + 1
        End If
        If chk_cytochemistry = 1 Then
            If ccount > 11 Then
                MsgBox "You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "U752"
            ccount = ccount + 1
        End If
''''''        If chk_sm = 1 Then
''''''            If ccount > 11 Then
''''''                MsgBox "You Can Select Max. of 11 Procedures only."
''''''                Unload Me
''''''                pathology.show
''''''                Exit Sub
''''''            End If
''''''            pathology.CMB_TEXT.AddItem "U753"
''''''            ccount = ccount + 1
''''''        End If
    '***********************************************
    ' Surface Marker single and complete added
        If chk_sm.Value = 1 And CHK_SMC.Value = 0 Then
        If ccount > 11 Then
            MsgBox "You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        pathology.CMB_TEXT.AddItem "U754"
        ccount = ccount + 1
    End If
    If CHK_SMC.Value = 1 And chk_sm.Value = 0 Then
        If ccount > 11 Then
            MsgBox "You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        pathology.CMB_TEXT.AddItem "U753"
        ccount = ccount + 1
    End If
    ' Surface Marker single and complete added
    '***********************************************
    
       If CHK_ANALYSIS.Value = 1 Then
            If CHK_CELLCOUNT.Value = 1 Then
                ind1 = CHK_CELLCOUNT.Caption
            Else
                ind1 = ""
            End If
            
            If chk_cytochemistry.Value = 1 Then
                ind2 = chk_cytochemistry.Caption
            Else
                ind2 = ""
            End If
            
            If chk_cytomorphology.Value = 1 Then
                ind3 = chk_cytomorphology.Caption
            Else
                ind3 = ""
            End If
            
            If CHK_BFC.Value = 1 Then
                ind4 = CHK_BFC.Caption
            Else
                ind4 = ""
            End If
            
            If chk_sm.Value = 1 Then
                ind5 = Trim(chk_sm.Caption)
            Else
                ind5 = ""
            End If
       End If
'End If
If bfmode = "NEW" Then
 If ccount = 0 Then
   MsgBox "No Service selected, You have to select atleast one service"
   Exit Sub
 End If
End If

Unload Me
pathology.show
End Sub
