VERSION 5.00
Begin VB.Form cardiologyind 
   Caption         =   "Indication"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5025
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
   LinkTopic       =   "Form2"
   ScaleHeight     =   5235
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Treatment Planned"
      Height          =   855
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   4695
      Begin VB.CheckBox chk_radiotherapy 
         Caption         =   "Radiotherapy"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   1695
      End
      Begin VB.CheckBox chk_chemotherapy 
         Caption         =   "Chemotherapy"
         Height          =   255
         Left            =   2160
         TabIndex        =   26
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox chk_Surgery 
         Caption         =   "Surgery"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmb_ok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Frame echoi 
      Caption         =   "2DECHO Indication"
      ForeColor       =   &H000000C0&
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   4695
      Begin VB.Frame Frame2 
         Height          =   1455
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   4095
         Begin VB.CheckBox CHK_2DIND9 
            Caption         =   "Post chemotherapy"
            Height          =   195
            Left            =   0
            TabIndex        =   44
            Top             =   0
            Width           =   2535
         End
         Begin VB.TextBox txt_drug 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   840
            TabIndex        =   38
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txt_dose 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2760
            TabIndex        =   37
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txt_noofcycles 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   840
            TabIndex        =   36
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txt_lastcycle 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   35
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox txt_bsa 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2760
            TabIndex        =   34
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Dose"
            Height          =   195
            Left            =   2160
            TabIndex        =   43
            Top             =   360
            Width           =   450
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Drug"
            Height          =   195
            Left            =   240
            TabIndex        =   42
            Top             =   360
            Width           =   420
         End
         Begin VB.Label Label3 
            Caption         =   "No.of Cycles"
            Height          =   375
            Left            =   120
            TabIndex        =   41
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Date of last cycle"
            Height          =   195
            Left            =   135
            TabIndex        =   40
            Top             =   1080
            Width           =   1515
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "BSA"
            Height          =   195
            Left            =   2235
            TabIndex        =   39
            Top             =   720
            Width           =   375
         End
      End
      Begin VB.CheckBox CHK_2DIND8 
         Caption         =   "ACT Clinic"
         Height          =   195
         Left            =   2160
         TabIndex        =   30
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox txt_others 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         TabIndex        =   28
         Top             =   2880
         Width           =   3495
      End
      Begin VB.CheckBox chk_2DIND7 
         Caption         =   "Chest Radiotherapy"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2640
         Width           =   2175
      End
      Begin VB.CheckBox chk_2DIND6 
         Caption         =   "Thoracic Surgery"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Width           =   1815
      End
      Begin VB.CheckBox CHK_2DIND5 
         Caption         =   "Pre chemotherapy"
         Height          =   255
         Left            =   2160
         TabIndex        =   5
         Top             =   720
         Width           =   2175
      End
      Begin VB.CheckBox CHK_2DIND4 
         Caption         =   "IHD"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox CHK_2DIND3 
         Caption         =   "Dyspnea on exertion/poor effort tolerance"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   4095
      End
      Begin VB.CheckBox CHK_2DIND2 
         Caption         =   "Cardiomegaly"
         Height          =   255
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
      Begin VB.CheckBox CHK_2DIND1 
         Caption         =   "ECG changes"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Others"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   3000
         Width           =   570
      End
   End
   Begin VB.Frame sti 
      Caption         =   "Stress Test Indication"
      ForeColor       =   &H000000C0&
      Height          =   3375
      Left            =   0
      TabIndex        =   17
      Top             =   1080
      Width           =   4695
      Begin VB.TextBox txt_stothers 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         TabIndex        =   32
         Top             =   2520
         Width           =   2775
      End
      Begin VB.CheckBox CHK_STIND6 
         Caption         =   "Pre-op Evaluation"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2160
         Width           =   3015
      End
      Begin VB.CheckBox CHK_STIND4 
         Caption         =   "K/c/o of IHD"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   3135
      End
      Begin VB.CheckBox CHK_STIND3 
         Caption         =   "Symptomatic for IHD with normal ECG"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   3855
      End
      Begin VB.CheckBox CHK_STIND2 
         Caption         =   "Symptomatic for IHD with ECG changes"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   3855
      End
      Begin VB.CheckBox CHK_STIND1 
         Caption         =   "Asymptomatic ECG changes"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   3375
      End
      Begin VB.CheckBox CHK_STIND5 
         Caption         =   "For functional capacity"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1800
         Width           =   3855
      End
      Begin VB.Label Label8 
         Caption         =   "Others"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   2520
         Width           =   615
      End
   End
   Begin VB.Frame pet 
      Caption         =   "PFT Indication"
      ForeColor       =   &H000040C0&
      Height          =   3375
      Left            =   0
      TabIndex        =   9
      Top             =   840
      Width           =   4695
      Begin VB.CheckBox CHK_PFIND4 
         Caption         =   "Pre operative"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1440
         Width           =   3255
      End
      Begin VB.CheckBox CHK_PFIND3 
         Caption         =   "Dyspnea for evaluation"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   3255
      End
      Begin VB.CheckBox CHK_PFIND2 
         Caption         =   "H/s/o of COPD"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   3015
      End
      Begin VB.CheckBox CHK_PFIND1 
         Caption         =   "H/s/o bronchial asthma"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   3015
      End
      Begin VB.CheckBox CHK_PFIND5 
         Caption         =   "Post operative"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Width           =   3375
      End
      Begin VB.CheckBox CHK_PFIND7 
         Caption         =   "Pre/Post Radiation"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2640
         Width           =   2895
      End
      Begin VB.CheckBox CHK_PFIND6 
         Caption         =   "Post chemotherapy"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   2280
         Width           =   3495
      End
   End
End
Attribute VB_Name = "cardiologyind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CHK_2DIND9_Click()
If CHK_2DIND9.Value = 1 Then
    txt_drug.Enabled = True
    txt_dose.Enabled = True
    txt_bsa.Enabled = True
    txt_noofcycles.Enabled = True
    txt_lastcycle.Enabled = True
Else
    txt_drug.Enabled = False
    txt_dose.Enabled = False
    txt_bsa.Enabled = False
    txt_noofcycles.Enabled = False
    txt_lastcycle.Enabled = False
End If
End Sub

Private Sub cmb_ok_Click()
CHKFLAG = False
'If ECNEWFLAG = True Or ECMODFLAG = True Or ecviewflag = True Then
    ind1 = ""
    ind2 = ""
    ind3 = ""
    ind4 = ""
    ind5 = ""
    ind6 = ""
    ind7 = ""
    ind8 = ""
    ind9 = ""
    tp1 = ""
    tp2 = ""
    tp3 = ""
    drug = ""
    dose = ""
    noofcycles = ""
    bsa = ""
    OTHERS = ""
    If chk_radiotherapy.Value = 1 Then
        tp1 = Trim(chk_radiotherapy.Caption)
    End If
    If chk_chemotherapy.Value = 1 Then
        tp1 = Trim(chk_chemotherapy.Caption)
    End If
    If chk_Surgery.Value = 1 Then
        tp1 = Trim(chk_Surgery.Caption)
    End If
    If pathology.txt_modcode.Text = "2D" Then
        If CHK_2DIND1.Value = 1 Then
             ind1 = Trim(CHK_2DIND1.Caption)
             CHKFLAG = True
        End If
        If CHK_2DIND2.Value = 1 Then
             ind2 = Trim(CHK_2DIND2.Caption)
             CHKFLAG = True
        End If
        If CHK_2DIND3.Value = 1 Then
             ind3 = Trim(CHK_2DIND3.Caption)
             CHKFLAG = True
        End If
        If CHK_2DIND4.Value = 1 Then
             ind4 = Trim(CHK_2DIND4.Caption)
             CHKFLAG = True
        End If
        If CHK_2DIND5.Value = 1 Then
             ind5 = Trim(CHK_2DIND5.Caption)
             CHKFLAG = True
        End If
        If chk_2DIND6.Value = 1 Then
             ind6 = Trim(chk_2DIND6.Caption)
             CHKFLAG = True
        End If
        If chk_2DIND7.Value = 1 Then
             ind7 = Trim(chk_2DIND7.Caption)
             CHKFLAG = True
        End If
        If CHK_2DIND8.Value = 1 Then
             ind8 = Trim(CHK_2DIND8.Caption)
             CHKFLAG = True
        End If
        If CHK_2DIND9.Value = 1 Then
             ind9 = Trim(CHK_2DIND9.Caption)
             CHKFLAG = True
        End If
        
        ind9 = ""
        drug = Trim(txt_drug.Text)
        dose = Trim(txt_dose.Text)
        noofcycles = Trim(txt_noofcycles.Text)
        bsa = Trim(txt_bsa.Text)
        OTHERS = Trim(TXT_OTHERS.Text)
        
    End If
    
    If pathology.txt_modcode.Text = "PF" Then
        If CHK_PFIND1.Value = 1 Then
            ind1 = Trim(CHK_PFIND1.Caption)
            CHKFLAG = True
        End If
        If CHK_PFIND2.Value = 1 Then
            ind2 = Trim(CHK_PFIND2.Caption)
            CHKFLAG = True
        End If
        If CHK_PFIND3.Value = 1 Then
            ind3 = Trim(CHK_PFIND3.Caption)
            CHKFLAG = True
        End If
        If CHK_PFIND4.Value = 1 Then
            ind4 = Trim(CHK_PFIND4.Caption)
            CHKFLAG = True
        End If
        If CHK_PFIND5.Value = 1 Then
            ind5 = Trim(CHK_PFIND5.Caption)
            CHKFLAG = True
        End If
        If CHK_PFIND6.Value = 1 Then
            ind6 = Trim(CHK_PFIND6.Caption)
            CHKFLAG = True
        End If
        If CHK_PFIND7.Value = 1 Then
            ind7 = Trim(CHK_PFIND7.Caption)
            CHKFLAG = True
        End If
        
        ind7 = ""
        ind8 = ""
        drug = ""
        dose = ""
        noofcycles = ""
        bsa = ""
        OTHERS = ""
    End If
    
    If pathology.txt_modcode.Text = "ST" Then
        If CHK_STIND1.Value = 1 Then
            ind1 = Trim(CHK_STIND1.Caption)
            CHKFLAG = True
        End If
         If CHK_STIND2.Value = 1 Then
            ind2 = Trim(CHK_STIND2.Caption)
            CHKFLAG = True
        End If
         If CHK_STIND3.Value = 1 Then
            ind3 = Trim(CHK_STIND3.Caption)
            CHKFLAG = True
        End If
         If CHK_STIND4.Value = 1 Then
            ind4 = Trim(CHK_STIND4.Caption)
            CHKFLAG = True
        End If
        
         If CHK_STIND5.Value = 1 Then
            ind5 = Trim(CHK_STIND5.Caption)
            CHKFLAG = True
        End If
         If CHK_STIND6.Value = 1 Then
            ind6 = Trim(CHK_STIND6.Caption)
            CHKFLAG = True
        End If
        ind6 = Trim(CHK_STIND6.Caption)
        ind7 = ""
        ind8 = ""
        drug = ""
        dose = ""
        noofcycles = ""
        bsa = ""
        OTHERS = Trim(txt_stothers)
        
    End If
 'End If
 
' If EDVIEWFLAG = True Then
'         ind1 = ""
'    ind2 = ""
'    ind3 = ""
'    ind4 = ""
'    ind5 = ""
'    ind6 = ""
'    ind7 = ""
'    ind8 = ""
'    ind9 = ""
'    tp1 = ""
'    tp2 = ""
'    tp3 = ""
'    drug = ""
'    dose = ""
'    noofcycles = ""
'    bsa = ""
'    others = ""
'    If chk_radiotherapy.Value = 1 Then
'        tp1 = Trim(chk_radiotherapy.Caption)
'    End If
'    If chk_chemotherapy.Value = 1 Then
'        tp1 = Trim(chk_chemotherapy.Caption)
'    End If
'    If chk_Surgery.Value = 1 Then
'        tp1 = Trim(chk_Surgery.Caption)
'    End If
'    If pathology.txt_modcode.Text = "2D" Then
'        If CHK_2DIND1.Value = 1 Then
'             ind1 = Trim(CHK_2DIND1.Caption)
'        End If
'        If CHK_2DIND2.Value = 1 Then
'             ind2 = Trim(CHK_2DIND2.Caption)
'        End If
'        If CHK_2DIND3.Value = 1 Then
'             ind3 = Trim(CHK_2DIND3.Caption)
'        End If
'        If CHK_2DIND4.Value = 1 Then
'             ind4 = Trim(CHK_2DIND4.Caption)
'        End If
'        If CHK_2DIND5.Value = 1 Then
'             ind5 = Trim(CHK_2DIND5.Caption)
'        End If
'        If chk_2DIND6.Value = 1 Then
'             ind6 = Trim(chk_2DIND6.Caption)
'        End If
'        If chk_2DIND7.Value = 1 Then
'             ind7 = Trim(chk_2DIND7.Caption)
'        End If
'        If CHK_2DIND8.Value = 1 Then
'             ind8 = Trim(CHK_2DIND8.Caption)
'        End If
'
'
'        ind9 = ""
'        drug = Trim(txt_drug.Text)
'        dose = Trim(txt_dose.Text)
'        noofcycles = Trim(txt_noofcycles.Text)
'        bsa = Trim(txt_bsa.Text)
'        others = Trim(txt_others.Text)
'
'    End If
'
'    If pathology.txt_modcode.Text = "PF" Then
'        If CHK_PFIND1.Value = 1 Then
'            ind1 = Trim(CHK_PFIND1.Caption)
'        End If
'        If CHK_PFIND2.Value = 1 Then
'            ind2 = Trim(CHK_PFIND2.Caption)
'        End If
'        If CHK_PFIND3.Value = 1 Then
'            ind3 = Trim(CHK_PFIND3.Caption)
'        End If
'        If CHK_PFIND4.Value = 1 Then
'            ind4 = Trim(CHK_PFIND4.Caption)
'        End If
'        If CHK_PFIND5.Value = 1 Then
'            ind5 = Trim(CHK_PFIND5.Caption)
'        End If
'        If CHK_PFIND6.Value = 1 Then
'            ind6 = Trim(CHK_PFIND6.Caption)
'        End If
'        If CHK_PFIND7.Value = 1 Then
'            ind7 = Trim(CHK_PFIND7.Caption)
'        End If
'
'        ind7 = ""
'        ind8 = ""
'        drug = ""
'        dose = ""
'        noofcycles = ""
'        bsa = ""
'        others = ""
'    End If
'
'    If pathology.txt_modcode.Text = "ST" Then
'        If CHK_STIND1.Value = 1 Then
'            ind1 = Trim(CHK_STIND1.Caption)
'        End If
'         If CHK_STIND2.Value = 1 Then
'            ind2 = Trim(CHK_STIND2.Caption)
'        End If
'         If CHK_STIND3.Value = 1 Then
'            ind3 = Trim(CHK_STIND3.Caption)
'        End If
'         If CHK_STIND4.Value = 1 Then
'            ind4 = Trim(CHK_STIND4.Caption)
'        End If
'
'         If CHK_STIND5.Value = 1 Then
'            ind5 = Trim(CHK_STIND5.Caption)
'        End If
'         If CHK_STIND6.Value = 1 Then
'            ind6 = Trim(CHK_STIND6.Caption)
'        End If
'        ind6 = Trim(CHK_STIND6.Caption)
'        ind7 = ""
'        ind8 = ""
'        drug = ""
'        dose = ""
'        noofcycles = ""
'        bsa = ""
'        others = Trim(txt_stothers)
'
'    End If
' End If
 If CHKFLAG = False Then
   i = MsgBox("You are not Selected any Indications, You Can select at least one", vbYesNo)
    If i = vbYes Then
        Exit Sub
    End If
 End If
    Unload Me
End Sub

Private Sub Form_Load()
    txt_drug.Enabled = False
    txt_dose.Enabled = False
    txt_bsa.Enabled = False
    txt_noofcycles.Enabled = False
    txt_lastcycle.Enabled = False
End Sub
