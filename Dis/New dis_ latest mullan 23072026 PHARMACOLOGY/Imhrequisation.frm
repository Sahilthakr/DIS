VERSION 5.00
Begin VB.Form IMHrequisation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IMMUNOHISTOCHEMISTRY"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
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
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.ComboBox CMB_TEXT 
         Height          =   315
         Left            =   720
         TabIndex        =   42
         Top             =   4920
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton CMD_NEXT 
         Caption         =   "&Next>>"
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
         Left            =   3960
         TabIndex        =   41
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CheckBox Check40 
         Caption         =   "Others"
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
         Height          =   495
         Left            =   4440
         TabIndex        =   40
         Top             =   4320
         Width           =   1215
      End
      Begin VB.CheckBox CHK_CKIT 
         Caption         =   "c-Kit"
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
         Height          =   495
         Left            =   4440
         TabIndex        =   39
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CheckBox CHK_SMA 
         Caption         =   "SMA"
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
         Height          =   495
         Left            =   4440
         TabIndex        =   38
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CheckBox CHK_MYGLOBIN 
         Caption         =   "Myglobin"
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
         Height          =   495
         Left            =   4440
         TabIndex        =   37
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CheckBox CHK_DESMIN 
         Caption         =   "Desmin"
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
         Height          =   495
         Left            =   4440
         TabIndex        =   36
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CheckBox CHK_FVIII 
         Caption         =   "F VIII"
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
         Height          =   495
         Left            =   2880
         TabIndex        =   35
         Top             =   4320
         Width           =   1215
      End
      Begin VB.CheckBox CHK_CD34 
         Caption         =   "CD34"
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
         Height          =   495
         Left            =   2880
         TabIndex        =   34
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CheckBox CHK_CD31 
         Caption         =   "CD31"
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
         Height          =   495
         Left            =   2880
         TabIndex        =   33
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CheckBox CHK_CD30 
         Caption         =   "CD30"
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
         Height          =   495
         Left            =   1320
         TabIndex        =   32
         Top             =   4320
         Width           =   1215
      End
      Begin VB.CheckBox CHK_CD20 
         Caption         =   "CD20"
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
         Height          =   495
         Left            =   1320
         TabIndex        =   31
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CheckBox CHK_CD15 
         Caption         =   "CD 15"
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
         Height          =   495
         Left            =   1320
         TabIndex        =   30
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CheckBox CHK_CD45 
         Caption         =   "CD 45 RO"
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
         Height          =   495
         Left            =   1320
         TabIndex        =   29
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CheckBox CHK_CD43 
         Caption         =   "CD43"
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
         Height          =   495
         Left            =   1320
         TabIndex        =   28
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CheckBox CHK_CD23 
         Caption         =   "CD23"
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
         Height          =   495
         Left            =   240
         TabIndex        =   27
         Top             =   4320
         Width           =   1215
      End
      Begin VB.CheckBox CHK_CD10 
         Caption         =   "CD10"
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
         Height          =   495
         Left            =   240
         TabIndex        =   26
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CheckBox CHK_CD3 
         Caption         =   "CD3"
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
         Height          =   495
         Left            =   240
         TabIndex        =   25
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CheckBox CHK_LCA 
         Caption         =   "LCA"
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
         Height          =   495
         Left            =   240
         TabIndex        =   24
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CheckBox CHK_MIC2 
         Caption         =   "MIC-2"
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
         Height          =   495
         Left            =   240
         TabIndex        =   23
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CheckBox CHK_NSE 
         Caption         =   "NSE"
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
         Height          =   495
         Left            =   4440
         TabIndex        =   22
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CheckBox CHK_SYNAPTOPHYSIN 
         Caption         =   "Synaptophysin"
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
         Height          =   495
         Left            =   4440
         TabIndex        =   21
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CheckBox CHK_CROMOGRANIN 
         Caption         =   "Chromogranin"
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
         Height          =   495
         Left            =   4440
         TabIndex        =   20
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CheckBox CHK_GFAB 
         Caption         =   "GFAP"
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
         Height          =   495
         Left            =   4440
         TabIndex        =   19
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox CHK_CALCITININ 
         Caption         =   "Calcitinin"
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
         Height          =   495
         Left            =   4440
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox CHK_THYROGLOBIN 
         Caption         =   "Thyroglobin"
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
         Height          =   495
         Left            =   2880
         TabIndex        =   17
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CheckBox CHK_S100P 
         Caption         =   "S 100 protein"
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
         Height          =   495
         Left            =   2880
         TabIndex        =   16
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CheckBox CHK_HMB45 
         Caption         =   "HMB 45"
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
         Height          =   495
         Left            =   1320
         TabIndex        =   15
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CheckBox CHK_PLAP 
         Caption         =   "PLAP"
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
         Height          =   495
         Left            =   2880
         TabIndex        =   14
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CheckBox CHK_AFETO 
         Caption         =   "a Feto Protein"
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
         Height          =   555
         Left            =   2880
         TabIndex        =   13
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox CHK_BHCG 
         Caption         =   "B HCG"
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
         Height          =   435
         Left            =   2880
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox CHK_CERB2 
         Caption         =   "c-erbB2"
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
         Height          =   435
         Left            =   1320
         TabIndex        =   11
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CheckBox CHK_PR 
         Caption         =   "PR"
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
         Height          =   435
         Left            =   1320
         TabIndex        =   10
         Top             =   1200
         Width           =   615
      End
      Begin VB.CheckBox CHK_ER 
         Caption         =   "ER"
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
         Height          =   495
         Left            =   1320
         TabIndex        =   9
         Top             =   720
         Width           =   615
      End
      Begin VB.CheckBox CHK_CK34BE12 
         Caption         =   "CK34BE12"
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
         Height          =   495
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox CHK_PSA 
         Caption         =   "PSA"
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
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   2040
         Width           =   735
      End
      Begin VB.CheckBox CHK_HBME1 
         Caption         =   "HBME 1"
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
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CheckBox CHK_CEA 
         Caption         =   "CEA"
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
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   735
      End
      Begin VB.CheckBox CHK_EMA 
         Caption         =   "EMA"
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
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   735
      End
      Begin VB.CheckBox CHK_CK 
         Caption         =   "CK"
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
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
      Begin VB.CheckBox CHK_VIM 
         Caption         =   "Vim"
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
         Height          =   495
         Left            =   2880
         TabIndex        =   2
         Top             =   3000
         Width           =   1455
      End
      Begin VB.CheckBox CHK_BCL 
         Caption         =   "Bcl-2"
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
         Height          =   495
         Left            =   2880
         TabIndex        =   1
         Top             =   2520
         Width           =   1095
      End
   End
End
Attribute VB_Name = "IMHrequisation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim INSDIHREPORT_QR As rdoQuery
Dim UPDDREQUISITNQR As rdoQuery




Private Sub CMD_NEXT_Click()
Dim ind_cbodept As Integer
On Error GoTo MSG:
If CHK_CK.Value = 1 Then
CMB_TEXT.AddItem "CK"
End If
If CHK_CK34BE12.Value = 1 Then
    CMB_TEXT.AddItem "CHK34BE12"
End If
If CHK_AFETO.Value = 1 Then
    CMB_TEXT.AddItem "A Feto Protein"

End If
If CHK_BCL.Value = 1 Then
    CMB_TEXT.AddItem "Bcl-2"
End If
If CHK_BHCG.Value = 1 Then
    CMB_TEXT.AddItem "B HCG"
End If
If CHK_CALCITININ.Value = 1 Then
    CMB_TEXT.AddItem "Calcitonin"
End If
If CHK_CD10.Value = 1 Then
    CMB_TEXT.AddItem "CD10"
End If
If CHK_CD15.Value = 1 Then
    CMB_TEXT.AddItem "CD15"
End If
If CHK_CD20.Value = 1 Then
    CMB_TEXT.AddItem "CD20"
End If
If CHK_CD23.Value = 1 Then
 CMB_TEXT.AddItem "CD23"
End If
If CHK_CD3.Value = 1 Then
    CMB_TEXT.AddItem "CD3"
End If
If CHK_CD30.Value = 1 Then
    CMB_TEXT.AddItem "CD30"
End If
If CHK_CD31.Value = 1 Then
    CMB_TEXT.AddItem "CD31"
End If
If CHK_CD34.Value = 1 Then
    CMB_TEXT.AddItem "CD34"
End If
If CHK_CD43.Value = 1 Then
    CMB_TEXT.AddItem "CD43"
End If
If CHK_CD45.Value = 1 Then
    CMB_TEXT.AddItem "CD 45 RO"
End If
If CHK_CEA.Value = 1 Then
    CMB_TEXT.AddItem "CEA"
End If
If CHK_CERB2.Value = 1 Then
    CMB_TEXT.AddItem "c-erbB2"
End If
If CHK_CKIT.Value = 1 Then
    CMB_TEXT.AddItem "C_KIT"
End If
If CHK_CROMOGRANIN.Value = 1 Then
CMB_TEXT.AddItem "Chromogranin"
End If
If CHK_DESMIN.Value = 1 Then
    CMB_TEXT.AddItem "Desmin"
End If
If CHK_EMA.Value = 1 Then
    CMB_TEXT.AddItem "EMA"
End If
If CHK_ER.Value = 1 Then
    CMB_TEXT.AddItem "ER"
End If
If CHK_FVIII.Value = 1 Then
    CMB_TEXT.AddItem "F VIII"
End If
If CHK_GFAB.Value = 1 Then
CMB_TEXT.AddItem "GFAP"
End If
If CHK_HBME1.Value = 1 Then
    CMB_TEXT.AddItem "HBME 1"
End If
If CHK_HMB45.Value = 1 Then
    CMB_TEXT.AddItem "HMB 45"
End If
If CHK_LCA.Value = 1 Then
    CMB_TEXT.AddItem "LCA"
End If
If CHK_MIC2.Value = 1 Then
CMB_TEXT.AddItem "MIC-2"
End If
If CHK_MYGLOBIN.Value = 1 Then
CMB_TEXT.AddItem "Myoglobin"
End If
If CHK_NSE.Value = 1 Then
CMB_TEXT.AddItem "NSE"
End If
If CHK_PLAP.Value = 1 Then
CMB_TEXT.AddItem "PLAP"
End If
If CHK_PR.Value = 1 Then
CMB_TEXT.AddItem "PR"
End If
If CHK_PSA.Value = 1 Then
CMB_TEXT.AddItem "PSA"
End If
If CHK_S100P.Value = 1 Then
CMB_TEXT.AddItem "S 100 protein"
End If
If CHK_SMA.Value = 1 Then
CMB_TEXT.AddItem "SMA"
End If
If CHK_SYNAPTOPHYSIN.Value = 1 Then
CMB_TEXT.AddItem "Synaptophysin"
End If
If CHK_THYROGLOBIN.Value = 1 Then
CMB_TEXT.AddItem "Thyroglobin"
End If
If CHK_VIM.Value = 1 Then
CMB_TEXT.AddItem "Vim"
End If
Dim str2 As String
str2 = ""
 '*** INSERTING INTO DIHREPORT
    If CMB_TEXT.ListCount > 0 Then
        For i = 0 To CMB_TEXT.ListCount - 1
            INSDIHREPORT_QR(0) = req_chglen(Trim(SPPATHREPORT.cmb_reqno.Text))
            INSDIHREPORT_QR(1) = CMB_TEXT.List(i)
            INSDIHREPORT_QR.Execute
            str2 = str2 & " " & CMB_TEXT.List(i)
        Next
        
        UPDDREQUISITNQR(0) = "YES"
        UPDDREQUISITNQR(1) = req_chglen(Trim(SPPATHREPORT.cmb_reqno.Text))
        UPDDREQUISITNQR.Execute
    End If

    '*** INSERTING INTO DIHREPORT
     SPPATHREPORT.txt_immhischerep.Text = Trim(str2)
     
'***** COMMENTED BY CHANDU 03/01/2003
''If pathology.CMB_TEXT.ListCount <= 5 Then
''pathology.IH
''
''   If Trim(pathology.cmb_ihprocode) = "" Then
''     For ind_cbodept = 0 To (pathology.cmb_ihprocode.ListCount - 1)
''        If Trim(pathology.cmb_ihprocode.List(ind_cbodept)) = "F501" Then
''            pathology.cmb_ihprocode.ListIndex = ind_cbodept
''        End If
''    Next ind_cbodept
''    End If
''
''
''
''
''End If
''If pathology.CMB_TEXT.ListCount > 5 Then
''pathology.cmb_ihprocode = ""
''pathology.IH
''   If Trim(pathology.cmb_ihprocode) = "" Then
''     For ind_cbodept = 0 To (pathology.cmb_ihprocode.ListCount - 1)
''        If Trim(pathology.cmb_ihprocode.List(ind_cbodept)) = "F502" Then
''            pathology.cmb_ihprocode.ListIndex = ind_cbodept
''        End If
''    Next ind_cbodept
''    End If
''
''End If

Unload Me
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub Form_Load()
    str = "INSERT INTO DISLIB.DIHREPORT(reqno,testorder) values (?,?)"
  Set INSDIHREPORT_QR = rdb.CreateQuery("insdihreport1_qrSP1", str)
  
  str = "update dislib.drequisitn set immhische=? where reqno=?"
  Set UPDDREQUISITNQR = rdb.CreateQuery("upddrequis123", str)
  
End Sub

Private Sub Form_Unload(cancel As Integer)
INSDIHREPORT_QR.Close
UPDDREQUISITNQR.Close
End Sub
