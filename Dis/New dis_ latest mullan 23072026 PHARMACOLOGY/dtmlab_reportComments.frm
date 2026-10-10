VERSION 5.00
Begin VB.Form dtmlab_reportComments 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7770
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Back"
      Height          =   495
      Left            =   6840
      TabIndex        =   6
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   495
      Left            =   6840
      TabIndex        =   5
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   1695
      Left            =   2640
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "dtmlab_reportComments.frx":0000
      Top             =   960
      Width           =   4095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00004000&
      BorderWidth     =   2
      Height          =   3015
      Left            =   0
      Top             =   0
      Width           =   7815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   6855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   0
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "dtmlab_reportComments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
