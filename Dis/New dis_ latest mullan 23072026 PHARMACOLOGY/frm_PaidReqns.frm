VERSION 5.00
Begin VB.Form frm_PaidReqns 
   Caption         =   "Existing Paid Requisitions Forms"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11445
   LinkTopic       =   "Form2"
   ScaleHeight     =   4440
   ScaleWidth      =   11445
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lbl_req 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DEPARTMENT SAMPLE COLLECTION WITH BARCODE "
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   345
      Left            =   1200
      TabIndex        =   1
      Top             =   840
      Width           =   8190
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   8595
   End
End
Attribute VB_Name = "frm_PaidReqns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MsgBox "Hi"
End Sub
