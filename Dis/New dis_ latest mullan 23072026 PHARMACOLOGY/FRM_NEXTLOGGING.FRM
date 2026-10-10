VERSION 5.00
Begin VB.Form frm_nextlogin 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1170
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   1995
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   FillColor       =   &H80000007&
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   ForeColor       =   &H8000000B&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   1995
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "Logging in..."
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frm_nextlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim i As Double
    DoEvents
    For i = 0 To 100
        i = i + 0.01
    Next i
    DIS_MAIN.show
End Sub
