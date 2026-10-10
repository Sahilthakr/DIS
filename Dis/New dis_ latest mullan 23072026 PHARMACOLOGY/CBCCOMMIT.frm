VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form cbccommit 
   Caption         =   "Form5"
   ClientHeight    =   10650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14475
   LinkTopic       =   "Form5"
   ScaleHeight     =   10650
   ScaleWidth      =   14475
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Case Wise Search"
      Height          =   315
      Left            =   11640
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txt_completed 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      ScrollBars      =   2  'Vertical
      TabIndex        =   69
      Top             =   11160
      Width           =   9015
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   240
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   1.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   240
      TabIndex        =   61
      Top             =   720
      Width           =   9015
      Begin VB.TextBox TXT_MINUTES 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1440
         TabIndex        =   63
         Top             =   120
         Width           =   495
      End
      Begin VB.TextBox TXT_OPERATOR 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         Left            =   4800
         TabIndex        =   62
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "TIME SLOT:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   79
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "OPERATED BY"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   78
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "MINUTES"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   77
         Top             =   240
         Width           =   855
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   8655
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "TIME SLOT:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   66
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "OPERATED BY"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   65
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "MINUTES"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   64
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton CMD_ERRLOGPRINT 
      BackColor       =   &H000000C0&
      Caption         =   "&ERROR LOG PRINT"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   11640
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   59
      Top             =   11640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5400
      TabIndex        =   58
      Top             =   11640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton CMD_START 
      BackColor       =   &H0080C0FF&
      Caption         =   "&START"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   11520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   1.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   11640
      Width           =   14775
      Begin VB.TextBox txt_srno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   360
         TabIndex        =   48
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txt_pathno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   47
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txt_reqno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   3360
         TabIndex        =   46
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox TXT_DCOMM 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   360
         TabIndex        =   45
         Top             =   2640
         Width           =   13935
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "HBsAg"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   360
         TabIndex        =   40
         Top             =   1320
         Width           =   2895
         Begin VB.TextBox txt_vhbsag 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            TabIndex        =   42
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txt_RHBsAg 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   1080
            TabIndex        =   41
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Value"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Result"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1080
            TabIndex        =   43
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "HCV"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3360
         TabIndex        =   35
         Top             =   1320
         Width           =   2895
         Begin VB.TextBox txt_VHCV 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   360
            TabIndex        =   37
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txt_RHCV 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   1200
            TabIndex        =   36
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Result"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1200
            TabIndex        =   39
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Value"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   360
            TabIndex        =   38
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "HIV"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   6360
         TabIndex        =   30
         Top             =   1320
         Width           =   3495
         Begin VB.TextBox txt_VHIV 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            TabIndex        =   32
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txt_RHIV 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   1440
            TabIndex        =   31
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Result"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1440
            TabIndex        =   34
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Value"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.TextBox TXT_CASENO 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   29
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Previous"
         Height          =   375
         Left            =   9960
         TabIndex        =   28
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Next"
         Height          =   375
         Left            =   9960
         TabIndex        =   27
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton CMD_MODIFY 
         Caption         =   "MODIFY"
         Height          =   255
         Left            =   5400
         TabIndex        =   26
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "SAVE"
         Height          =   375
         Left            =   11160
         TabIndex        =   25
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label12 
         BackColor       =   &H008080FF&
         Caption         =   "Label11"
         Height          =   135
         Left            =   5400
         TabIndex        =   56
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Label8"
         Height          =   255
         Left            =   5400
         TabIndex        =   55
         Top             =   840
         Width           =   2175
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00FFFFC0&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   3015
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   360
         Width           =   14535
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "SRNO"
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
         Left            =   360
         TabIndex        =   54
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "MICRO NO."
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
         Left            =   1080
         TabIndex        =   53
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "REQ NO"
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
         Left            =   3360
         TabIndex        =   52
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "COMMENT"
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
         Left            =   360
         TabIndex        =   51
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "CASENO"
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
         Left            =   2160
         TabIndex        =   50
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label LBL_NAME 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   225
         Left            =   5400
         TabIndex        =   49
         Top             =   720
         Width           =   45
      End
   End
   Begin VB.CommandButton cmd_back 
      BackColor       =   &H0080C0FF&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   10080
      Width           =   1935
   End
   Begin VB.CommandButton CMD_EXIT 
      BackColor       =   &H0080C0FF&
      Caption         =   "&EXIT"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   10080
      Width           =   1935
   End
   Begin VB.ComboBox TXT_COMMENTS1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "CBCCOMMIT.frx":0000
      Left            =   480
      List            =   "CBCCOMMIT.frx":0002
      TabIndex        =   21
      Top             =   11640
      Visible         =   0   'False
      Width           =   13935
   End
   Begin VB.CommandButton CMD_REFRESH 
      BackColor       =   &H0080C0FF&
      Caption         =   "&REFRESH"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   10080
      Width           =   1935
   End
   Begin VB.TextBox TXT_MESSAGE 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   360
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   11160
      Width           =   9015
   End
   Begin VB.ComboBox cmb_srno 
      Height          =   315
      Left            =   9840
      TabIndex        =   18
      Top             =   1800
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   1.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9255
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   13935
      Begin VB.TextBox TXT_DATA 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   8040
         TabIndex        =   9
         Top             =   1800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox TXT_COMMENTS 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "CBCCOMMIT.frx":0004
         Left            =   3840
         List            =   "CBCCOMMIT.frx":0006
         TabIndex        =   8
         Top             =   3480
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.ComboBox cmb_modality 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "CBCCOMMIT.frx":0008
         Left            =   2880
         List            =   "CBCCOMMIT.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   120
         Width           =   6615
      End
      Begin VB.CommandButton cmd_showrpt 
         BackColor       =   &H000080FF&
         Caption         =   "&Show Provisional Data"
         Height          =   315
         Left            =   7680
         TabIndex        =   5
         Top             =   480
         Width           =   1815
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   2400
         TabIndex        =   11
         Top             =   840
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSMask.MaskEdBox txt_frm 
         Height          =   315
         Left            =   3840
         TabIndex        =   12
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt_to 
         Height          =   315
         Left            =   6000
         TabIndex        =   13
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cmb_dept 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8160
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   3240
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   6735
         Left            =   600
         TabIndex        =   7
         Top             =   1680
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   11880
         _Version        =   393216
         Rows            =   20
         Cols            =   8
         MergeCells      =   1
         Appearance      =   0
         FormatString    =   $"CBCCOMMIT.frx":000C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lbl_modality 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Modality  :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   240
         Left            =   840
         TabIndex        =   17
         Top             =   120
         Width           =   1860
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Date  :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   240
         Left            =   960
         TabIndex        =   16
         Top             =   480
         Width           =   1440
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   15
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5280
         TabIndex        =   14
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.CommandButton cmd_pro_data 
      Caption         =   "Show Commited Data"
      Height          =   315
      Left            =   9840
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   2880
      Width           =   2535
   End
   Begin VB.CommandButton cmd_fullreport 
      BackColor       =   &H000000FF&
      Caption         =   "EXIT"
      Height          =   495
      Left            =   10440
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   0
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   0
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSMask.MaskEdBox msk_sysdate 
      Height          =   315
      Left            =   9960
      TabIndex        =   67
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox msk_systime 
      Height          =   315
      Left            =   11520
      TabIndex        =   68
      Top             =   600
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "hh:mm:ss"
      Mask            =   "##:##:##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9360
      TabIndex        =   76
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "MESSAGE"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12240
      TabIndex        =   75
      Top             =   11520
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ENDTIME"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   74
      Top             =   11640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TIMER"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   73
      Top             =   11640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lbl_microbiology 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MICROBIOLOGY AUTHENTICATION SCREEN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   525
      Left            =   2280
      TabIndex        =   72
      Top             =   0
      Width           =   9615
   End
   Begin VB.Label LBL_PHASE 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1200
      TabIndex        =   71
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   495
      Left            =   7080
      TabIndex        =   70
      Top             =   3000
      Width           =   1215
   End
End
Attribute VB_Name = "cbccommit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public req_no_tst As String
Dim sel_moadlityqr  As rdo.rdoQuery
Dim sel_deptqr     As rdo.rdoQuery
 Dim SELDMUL_QR As rdoQuery
Dim SELDMUL_RS As rdoResultset
Public C1 As String
Dim sel_modalityrs      As rdo.rdoResultset
Dim sel_deptrs          As rdo.rdoResultset
Dim SEL_MODALITY_RS     As rdo.rdoResultset
Dim PATDETAILS As String
Dim txt_modcode As String
Dim MODPHASE As Boolean
Public DATE11 As String
Public click As Integer
Public time1 As String
Public VAR_REQNO As Variant
Public micro_no As String
Dim ipadd As String
Dim srno As Long
Dim obj As HMSCrypt
Dim key As Integer
Dim key1 As String
Dim INSGRID As Boolean
Dim PATNAME As String
Dim pname As String
Dim wbval As String
Dim pno As String
Dim sac As String
'''''SELECT QUERIES''''''''
Dim flag As String
Dim rs_minterface As rdoResultset
Dim qr_minterface As rdoQuery
Dim mod_reqno As String
Dim SELDHEMCYTO_QR As rdoQuery
Dim SELDHEMCYTO_RS As rdoResultset
Dim SELCBCINTFERR_QR As rdoQuery
Dim SELCBCINTFERR_QR_sp As rdoQuery
Dim SELCBCINTFERR_RS As rdoResultset
Dim SELCBCINTFERR_RS_sp As rdoResultset
Dim USRNAME_QR As rdoQuery
Dim USRNAME_RS As rdoResultset
Dim SELOPRTRQR As rdoQuery
Dim SELOPRTRRS As rdoResultset
Dim selhintfqr As rdoQuery
Dim selhintfqr1 As rdoQuery
Dim selhintfrs As rdoResultset
Dim SELSEXAGE_QR As rdoQuery
Dim SELSEXAGE_RS As rdoResultset
Dim SELRREQUISITN_QR As rdoQuery
Dim SELRREQUISITN_RS As rdoResultset
Dim selta1012_qr As rdoQuery
Dim selta1012_rs As rdoResultset
Dim SELDHEMCBCQR As rdoQuery
Dim SELDHEMCBCRS As rdoResultset
Dim SELCASENO_QR As rdoQuery
Dim SELCASENO_RS As rdoResultset
Dim SELTIMESLOT_QR As rdoQuery
Dim SELTIMESLOT_RS As rdoResultset
Dim SELHINTFVAL_QR As rdoQuery
Dim SELHINTFVAL_RS As rdoResultset
Dim SELPATNSTAT_QR As rdoQuery
Dim SELPATNSTAT_RS As rdoResultset
Dim SELSTAFF_QR As rdoQuery
Dim SELSTAFF_RS As rdoResultset

Dim SELMICROLIST_QR As rdoQuery
Dim SELMICROLIST_RS As rdoResultset

''''END OF SELECT QUERIES''''''
    Public req As String
Public case_no_tst As String
''''INSERT AND UPDATE QUERIES''''
Dim ins_micro As rdoQuery
Dim ins_micro_temp As rdoQuery
Dim dlt_micro As rdoQuery

Dim phase As String
Dim INSCBCINTFERR_QR As rdoQuery
Dim INSDHEMCBCDET_QR As rdoQuery
Dim INSMICROLIST_QR As rdoQuery
Dim DELMICROLIST_QR As rdoQuery
Dim updhinterface_qr As rdoQuery
Dim updmtemp_qr As rdoQuery
Dim updmproce_qr As rdoQuery
Dim updrrequisitn1_qr As rdoQuery
Dim UPDRTIMEC_QR As rdoQuery
Dim UPDCBCINTFERR_QR As rdoQuery
Dim UPDHINTFQR As rdoQuery
Dim updrrequisitnt_qr As rdoQuery
Dim ins_dcbcrep1_qr As rdoQuery
''''END OF INSERT & UPDATE QUERIES''''

Dim FILLERROR As Boolean
Dim SELDHEMCBCQRFLAG As Boolean
Dim NTPHLABSFLAG As Boolean
Dim LMPCYTEABSFLAG As Boolean
Dim MNCYTEABSFLAG As Boolean
Dim ESNPHLABSFLAG As Boolean
Dim BSPHLABSFLAG As Boolean
Dim ERROR As Boolean
Dim SAVEERROR As Boolean
Dim ROUTINE As Boolean
Dim ENDOFPROCESS As Boolean

Dim LOGINUSRNAME As String
Dim LOGINTYP As String
Dim LOGINUSRID As String
Dim str As String
Dim HAGE As String
Dim HBSEX As String
Dim HBHV As String
Dim HBLV As String
Dim range
Dim SA, age, SEX As String
Dim h, m, dur As String
Dim tdiff, I1, ddiff As Double
Dim I, I2 As Long
''''''''
Dim RESULTQR As rdoQuery
Dim RESULTRS As rdoResultset
Dim resultflag As Boolean

'***** Shraddha 27/09/10 *****
Dim SEL_MICROPROCE_QR As rdoQuery
Dim SEL_MICROPROCE_RS As rdoResultset
Dim SEL_MICROTEMP_QR As rdoQuery
Dim SEL_MICROTEMP_RS As rdoResultset

Dim SEL_DMBCCOMM_QR As rdoQuery
Dim SEL_DMBCCOMM_RS As rdoResultset
Dim UPD_DMBCCOMM_qr As rdoQuery
Dim UPD_MINTERFACE_qr As rdoQuery


Private Sub cmb_modality_Change()
 txt_frm.Text = Format(DATE11, "DD/MM/YYYY")
  txt_to.Text = Format(DATE11, "DD/MM/YYYY")
End Sub

Private Sub cmd_back_Click()
DIS_MAIN.show
cbccommit.Hide
End Sub

'Private Sub CHK_Save_Click()
'If CHK_Save.Value = 1 Then
'MSFlexGrid1.col = 7
'MSFlexGrid1.ROW = MSFlexGrid1.ROW
'MSFlexGrid1.Text = CHK_Save.Caption
'    MSFlexGrid1.Text = "No"
'    MSFlexGrid1.col = 7
'    MSFlexGrid1.CellBackColor = vbWhite
'    'CHK_Save.Visible = False
'Else
'    MSFlexGrid1.Text = "Yes"
''        MSFlexGrid1.ROW = I
'        MSFlexGrid1.col = 7
'        MSFlexGrid1.CellBackColor = &HC0C0C0
'        CHK_Save.Visible = False
'End If
''CHK_Save.Visible = False
'End Sub

Private Sub CMD_ERRLOGPRINT_Click()
Dim d As String
d = Format(Date, "DD-MM-YYYY")
CrystalReport2.ReportFileName = "C:\DISREPORTSEC\CBCERRREP.rpt"
CrystalReport2.SelectionFormula = "{CBCINTFERR.ERRORDATE}=DATE('" & d & "') AND {CBCINTFERR.PRINTFLAG}<>'P'"
CrystalReport2.WindowState = crptMaximized
CrystalReport2.Destination = crptToPrinter
CrystalReport2.Action = 1
cmd_exit.Enabled = True

'UPDATING INTERFACE.CBCINTFERR FOR PRINTFLAG ='P'
SELCBCINTFERR_QR(0) = d
SELCBCINTFERR_QR(1) = "N"
Set SELCBCINTFERR_RS = SELCBCINTFERR_QR.OpenResultset(1, 1)

If SELCBCINTFERR_RS.RowCount > 0 Then

    While SELCBCINTFERR_RS.EOF = False

        UPDCBCINTFERR_QR(0) = "P"
        UPDCBCINTFERR_QR(1) = SELCBCINTFERR_RS(0)
        UPDCBCINTFERR_QR(2) = SELCBCINTFERR_RS(1)
        UPDCBCINTFERR_QR(3) = SELCBCINTFERR_RS(2)
        UPDCBCINTFERR_QR(4) = SELCBCINTFERR_RS(4)
        UPDCBCINTFERR_QR(5) = SELCBCINTFERR_RS(5)
        UPDCBCINTFERR_QR.Execute

        SELCBCINTFERR_RS.MoveNext

    Wend
End If
End Sub

Private Sub cmd_exit_Click()
End 'SHRADDHA
If FILLERROR = False And SAVEERROR = False And PRINTERROR = False Then
    ENDOFPROCESS = True

Else
    ENDOFPROCESS = False
End If

End Sub

Private Sub cmd_fullreport_Click()
End
End Sub

Private Sub cmd_modify_Click()
'MODPHASE = True
'txt_vhbsag.Enabled = False
'txt_RHBsAg.Enabled = True
'txt_VHCV.Enabled = False
'txt_RHCV.Enabled = True
'txt_VHIV.Enabled = False
'txt_RHIV.Enabled = True
'TXT_DCOMM.Enabled = True
'TXT_COMMENTS1.Visible = True
'txt_RHBsAg.SetFocus
End Sub

Private Sub cmd_Ok_Click()
MousePointer = vbHourglass
On Error GoTo exet1
Dim K As Integer
K = 0
'level = "SAVE"
For I = 1 To MSFlexGrid1.Rows - 1
rdb.BeginTrans
If UCase(Trim(MSFlexGrid1.TextMatrix(I, 7))) = "NO" Then
    commit = False
    flag = "NO"
Else
    flag = "YES"
    K = K + 1
        If MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 8) = "URGENT" Then
        flag = "URG"

    End If

    'micro
    commit = True
    rno = req_chglen(Trim(MSFlexGrid1.TextMatrix(I, 3)))
    UPDATE

End If
If MSFlexGrid1.Rows > 1 Then
  'INSERTTOLISTTBL
End If
rdb.CommitTrans

txt_completed = I + 1 & " Record Completed"
Next
If K > 1 Then
MsgBox "RECORD SUCCESSFULLY SAVED"
End If
'  CHK_Save.Top = 720
'  CHK_Save.Left = 240
'  CHK_Save.Value = 0
'  'MSFlexGrid1.Clear
'  CHK_Save.Visible = False

  FILLGRID
  Exit Sub

exet1:

  MsgBox Err.Description
rdb.RollbackTrans
   INSCBCINTFERR_QR(0) = "GZZSE" & req_chglen(Trim(MSFlexGrid1.TextMatrix(I, 3)))
    INSCBCINTFERR_QR(1) = Err.Number
    INSCBCINTFERR_QR(2) = "NEWEXE"
    INSCBCINTFERR_QR(3) = Err.Description
    INSCBCINTFERR_QR(4) = Date
    INSCBCINTFERR_QR(5) = Time
    INSCBCINTFERR_QR(6) = "N"
    INSCBCINTFERR_QR.Execute
    ERROR = True


  MousePointer = vbNormal
End Sub
Sub INSERTTOLISTTBL()
SELMICROLIST_QR(0) = "GZZSE" & Trim(MSFlexGrid1.TextMatrix(I, 3))
Set SELMICROLIST_RS = SELMICROLIST_QR.OpenResultset(1, 1)
If SELMICROLIST_RS.RowCount > 0 Then
DELMICROLIST_QR(0) = "GZZSE" & Trim(MSFlexGrid1.TextMatrix(I, 3))
DELMICROLIST_QR.Execute
End If

INSMICROLIST_QR(0) = Trim(msk_sysdate.FormattedText)
INSMICROLIST_QR(1) = Trim(MSFlexGrid1.TextMatrix(I, 2))
INSMICROLIST_QR(2) = "GZZSE" & Trim(MSFlexGrid1.TextMatrix(I, 3))
INSMICROLIST_QR(3) = Trim(MSFlexGrid1.TextMatrix(I, 1))
INSMICROLIST_QR(4) = GETNAME(Trim(MSFlexGrid1.TextMatrix(I, 2)))
INSMICROLIST_QR(5) = Trim(MSFlexGrid1.TextMatrix(I, 4))
INSMICROLIST_QR(6) = Trim(MSFlexGrid1.TextMatrix(I, 5))
INSMICROLIST_QR(7) = Trim(MSFlexGrid1.TextMatrix(I, 6))
INSMICROLIST_QR(8) = Trim(MSFlexGrid1.TextMatrix(I, 7))
INSMICROLIST_QR(9) = Trim(MSFlexGrid1.TextMatrix(I, 8))
INSMICROLIST_QR(10) = Trim(MSFlexGrid1.TextMatrix(I, 9))
INSMICROLIST_QR(11) = flag
INSMICROLIST_QR(12) = Trim(msk_sysdate.FormattedText)
INSMICROLIST_QR(13) = Trim(msk_systime.FormattedText)
INSMICROLIST_QR(14) = Trim(TXT_OPERATOR.Text)
INSMICROLIST_QR(15) = Winsock1.LocalIP
INSMICROLIST_QR(16) = "1"


SEL_DMBCCOMM_QR(0) = "GZZSE" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3))
Set SEL_DMBCCOMM_RS = SEL_DMBCCOMM_QR.OpenResultset(2, 1)
If Not SEL_DMBCCOMM_RS.EOF Then
'Else
On Error Resume Next

COMM = Trim(SEL_DMBCCOMM_RS(1))

End If

INSMICROLIST_QR(17) = Trim(COMM)
INSMICROLIST_QR.Execute

End Sub
Function GETNAME(ByVal CASENO As String) As String
SELSEXAGE_QR(0) = CASENO
Set SELSEXAGE_RS = SELSEXAGE_QR.OpenResultset(1, 1)
If SELSEXAGE_RS.RowCount > 0 Then
PATNAME = Trim(SELSEXAGE_RS("TAG")) & " " & Trim(SELSEXAGE_RS("FIRST_NAME")) & " " & Trim(SELSEXAGE_RS("LAST_NAME"))
Else
PATNAME = ""
End If
GETNAME = PATNAME
End Function

Private Sub cmd_pro_data_Click()
Text3.Text = "Commited Reports."
click = "2"
show_data
End Sub

Private Sub cmd_refresh_Click()
srno = srno '+ 1
MSFlexGrid1.clear
'FILLGRID
End Sub

Public Sub cmd_showrpt_Click()
click = "1"
show_data
Text3.Text = "Provision Reports."
End Sub
Public Function show_data()
req_no_tst = ""
case_no_tst = ""
cmb_srno.clear
ProgressBar1.Visible = True
MSFlexGrid1.clear
ProgressBar1.Visible = True 'Shraddha Pandit 17/07/09
ProgressBar1.Min = 0
ProgressBar1.Max = 100
ProgressBar1.Value = 0

While ProgressBar1.Value < 45
    MYDELAY (1)

    ProgressBar1.Value = ProgressBar1.Value + 45
Wend

FILLGRID
 ProgressBar1.Value = 100
ProgressBar1.Visible = False
End Function
Private Sub Command1_Click()
If txt_srno.Text = 1 Then
Else
MSFlexGrid1.row = Val(txt_srno.Text) - 1
End If
DISPLAYINTXT (MSFlexGrid1.row)
End Sub

Private Sub Command2_Click()
If txt_srno.Text + 1 = MSFlexGrid1.Rows Then
Else
MSFlexGrid1.row = Val(txt_srno.Text) + 1
End If
DISPLAYINTXT (MSFlexGrid1.row)
End Sub
Sub PatInfo(ByVal CASENO As String, ByVal REQNO As String, ByVal patstat As String)
Dim patcat As String

Dim micro As String
Dim ward As String
Dim bed As String

patcat = ""
If case_no_tst <> "" Then
    CASENO = case_no_tst
       Else
       CASENO = CASENO
    End If

If Len(Trim(CASENO)) = 6 Then
patcat = "S"
ElseIf Len(Trim(CASENO)) = 8 Then
patcat = "P"
End If

Dim SELbas_QR As rdoQuery
Dim SELbas_RS As rdoResultset
Dim procode As String
If Trim(patcat) = "S" Then
    str = "SELECT * FROM dislib.dreqstaffd where reqno=?"
    Set SELbas_QR = rdb.CreateQuery("", str)
    If req_no_tst <> "" Then
       SELbas_QR(0) = req_no_tst
       Else
         SELbas_QR(0) = REQNO
    End If

    Set SELbas_RS = SELbas_QR.OpenResultset(1, 1)
    If SELbas_RS.EOF = False Then
'    procode = Trim(SELbas_RS(0))
    pname = Trim(SELbas_RS("PNAME"))
    PATNAME = pname
    sac = Trim(SELbas_RS("SEX")) & " / " & Trim(SELbas_RS("AGE")) '& "ST"
    End If
    PATDETAILS = pname
Else

    str = "SELECT * FROM patoolslib.ta101 where case_no=?"
    Set SELbas_QR = rdb.CreateQuery("", str)
    If case_no_tst <> "" Then
       SELbas_QR(0) = case_no_tst
       Else
        SELbas_QR(0) = CASENO
    End If

    Set SELbas_RS = SELbas_QR.OpenResultset(1, 1)
    If SELbas_RS.EOF = False Then
'    procode = Trim(SELbas_RS(0))
            pname = Trim(SELbas_RS("TAG")) & " " & Trim(SELbas_RS("FIRST_NAME")) & " " & Trim(SELbas_RS("LAST_NAME"))
            PATNAME = pname
            sac = Trim(SELbas_RS("SEX")) & " / " & DateDiff("yyyy", SELbas_RS("BIRTH_DATE"), Trim(msk_sysdate.FormattedText)) & ""
    End If
    PATDETAILS = pname

End If







End Sub


Sub ADDTOGRID()
Dim RW As Integer
RW = txt_srno.Text

MSFlexGrid1.TextMatrix(RW, 0) = Trim(txt_srno.Text)
MSFlexGrid1.TextMatrix(RW, 1) = Trim(txt_pathno.Text)
MSFlexGrid1.TextMatrix(RW, 2) = Trim(TXT_CASENO.Text)
MSFlexGrid1.TextMatrix(RW, 3) = Trim(txt_vhbsag.Text)
MSFlexGrid1.TextMatrix(RW, 4) = Trim(txt_vhbsag.Text)
MSFlexGrid1.TextMatrix(RW, 5) = Trim(txt_RHBsAg.Text)
MSFlexGrid1.TextMatrix(RW, 6) = Trim(txt_VHCV.Text)
MSFlexGrid1.TextMatrix(RW, 7) = Trim(txt_reqno.Text)
MSFlexGrid1.TextMatrix(RW, 8) = Trim(txt_VHIV.Text)
MSFlexGrid1.TextMatrix(RW, 9) = Trim(txt_RHIV.Text)
MSFlexGrid1.TextMatrix(RW, 10) = Trim(TXT_DCOMM.Text)
'MSFlexGrid1.TextMatrix(RW, 11) = ""

 key1 = makekey
'UPDATE IN TABLES
I = RW
UPDATE

'STR = "UPDATE DISLIB.MICROPROCE SET OBSERVALUE=?,RESULT=?,NRANGE=?,FLD2=?,FLD1=? WHERE REQNO=? AND PRO_CODE=?"
 '   Set updmproce_qr = rdb.CreateQuery("", STR)

  '  STR = "UPDATE DISLIB.MICROTEMP SET OBSERVALUE=?,RESULT=? ,FLD1=? WHERE REQNO=? AND PRO_CODE=?"
  '  Set updmtemp_qr = rdb.CreateQuery("", STR)

updmtemp_qr(0) = Trim(txt_vhbsag.Text)
updmtemp_qr(1) = Trim(txt_RHBsAg.Text)
updmtemp_qr(2) = key1
updmtemp_qr(3) = "GZZSE" & Trim(txt_reqno.Text)
updmtemp_qr(4) = "G129"
updmtemp_qr.Execute


updmtemp_qr(0) = Trim(txt_VHCV.Text)
updmtemp_qr(1) = Trim(txt_RHCV.Text)
updmtemp_qr(2) = key1
updmtemp_qr(3) = "GZZSE" & Trim(txt_reqno.Text)
updmtemp_qr(4) = "G134"
updmtemp_qr.Execute


updmtemp_qr(0) = Trim(txt_VHIV.Text)
updmtemp_qr(1) = Trim(txt_RHIV.Text)
updmtemp_qr(2) = key1
updmtemp_qr(3) = "GZZSE" & Trim(txt_reqno.Text)
updmtemp_qr(4) = "G171"
updmtemp_qr.Execute

'------micro proce


 'OBSERVALUE=?,RESULT=?,NRANGE=?,FLD2=?,FLD1=? WHERE REQNO=? AND PRO_CODE=?"
updmproce_qr(0) = obj.Encrypt(Trim(txt_vhbsag.Text), key)
updmproce_qr(1) = obj.Encrypt(Trim(txt_RHBsAg.Text), key)
updmproce_qr(2) = Trim(txt_RHBsAg.Text)
updmproce_qr(3) = Trim(txt_vhbsag.Text)
updmproce_qr(4) = key1
updmproce_qr(5) = "GZZSE" & Trim(txt_reqno.Text)
updmproce_qr(6) = "G129"
updmproce_qr.Execute


updmproce_qr(0) = obj.Encrypt(Trim(txt_VHCV.Text), key)
updmproce_qr(1) = obj.Encrypt(Trim(txt_RHCV.Text), key)
updmproce_qr(2) = Trim(txt_RHCV.Text)
updmproce_qr(3) = Trim(txt_VHCV.Text)
updmproce_qr(4) = key1
updmproce_qr(5) = "GZZSE" & Trim(txt_reqno.Text)
updmproce_qr(6) = "G134"
updmproce_qr.Execute


updmproce_qr(0) = obj.Encrypt(Trim(txt_VHIV.Text), key)
updmproce_qr(1) = obj.Encrypt(Trim(txt_RHIV.Text), key)
updmproce_qr(2) = Trim(txt_RHIV.Text)
updmproce_qr(3) = Trim(txt_VHIV.Text)
updmproce_qr(4) = key1
updmproce_qr(5) = "GZZSE" & Trim(txt_reqno.Text)
updmproce_qr(6) = "G171"
updmproce_qr.Execute

End Sub

Private Sub Command3_Click()
ADDTOGRID
    'FLAG = "YES"

On Error GoTo NX
    rdb.BeginTrans
    micro

    'COMMIT = True
    'RNO = "GZZSE" & req_chglen(Trim(MSFlexGrid1.TextMatrix(I, 2)))
   ' UPDATE
   rdb.CommitTrans
   Exit Sub
NX:
rdb.RollbackTrans
   MsgBox Err.Description
End Sub

Private Sub Command4_Click()
 Report_microbiology1.show
  'Report_editmicrobiology1.show
  cbccommit.Hide
     Report_microbiology1.cmd_cancel_Click
       Report_microbiology1.CMD_NEW.Enabled = False
         Report_microbiology1.cmd_view_Click
        ' Report_microbiology1.VIEW_FLAG.Value = True
         VIEW_FLAG = True

         Report_microbiology1.MaskEdBox1.Text = ""
            Report_microbiology1.cmb_reqno.Text = ""
            Report_microbiology1.MaskEdBox1.SetFocus
              Report_microbiology1.Opt_patient.Value = True
End Sub

Private Sub Form_Activate()
'CD/23156
Dim DEL_QR As rdoQuery
'************* Shraddha ***********
Dim SEL_INT_QR As rdoQuery
Dim SEL_INT_RS As rdoResultset
Dim SQL_STR As String

Dim SEL_CHK_QR As rdoQuery
Dim SEL_CHK_RS As rdoResultset
Dim SQL_CHK_STR As String

'**********************************
'Set DEL_QR = rdb.CreateQuery("", "DELETE FROM DISLIB.MICROTEMP")

    ROUTINE = False

  '  FILLGRID

    'DEL_QR.Execute
'    Do While ROUTINE = False
'            If I = 1 Then
'                TXT_MESSAGE.Text = "  NO RECORDS FOUND WITHIN GIVEN TIME SLOT ....."
'                GoTo retry
'            End If
'            TXT_MESSAGE.Text = "  AUTOMATION PROCESS STARTED .....!!!!!!"
'
'            For I2 = 1 To MSFlexGrid1.Rows - 1
'                I = 1
'                'MICRO 'INSERTING INTO MICROPROC
'
'                If SAVEERROR = False Then
'                    'UPDATE    'UPDATING MINTERFACE, RREQUISITN & RTIMELOG
'                    MSFlexGrid1.TextMatrix(I, 4) = "COMMITTED"
'                    MYDELAY (5)
'                    If MSFlexGrid1.Rows > 1 Then
'
'                    '*********** Shraddha
'
'                        R_NO = req_chglen(Trim(MSFlexGrid1.TextMatrix(I, 2)))
'                        SQL_STR = "SELECT * FROM INTERFACE.MINTERFACE WHERE REQNO='" & R_NO & "' and IFLAG='2' AND TESTRESULT <> '' AND FLD2<>''"
'                        Set SEL_INT_QR = rdb.CreateQuery("", SQL_STR) '''added by charulata
'
'
'
'                        SQL_CHK_STR = "SELECT * FROM INTERFACE.MINTERFACE WHERE REQNO='" & R_NO & "'"
'                        Set SEL_CHK_QR = rdb.CreateQuery("", SQL_CHK_STR)
'
'
'
'                        Set SEL_INT_RS = SEL_INT_QR.OpenResultset(1, 1)
'                        Set SEL_CHK_RS = SEL_CHK_QR.OpenResultset(1, 1)
'
'
'                        If SEL_INT_RS.RowCount = SEL_CHK_RS.RowCount Then
'                            MICRO 'INSERTING INTO MICROPROC
'                            UPDATE    'UPDATING MINTERFACE, RREQUISITN & RTIMELOG
'                            PRINTING
'
''                        End If
''                    End If
'                            If ERROR = False Then
'                                MSFlexGrid1.TextMatrix(I, 4) = "PRINTED"
'                                MYDELAY (5)
'                            Else
'                                MSFlexGrid1.TextMatrix(I, 4) = "PRINT ERROR"
'                                MYDELAY (5)
'                                If MSFlexGrid1.Rows <> 2 Then
'                                    MSFlexGrid1.RemoveItem I
'                                Else
'                                    MSFlexGrid1.Clear
'                                    MSFlexGrid1.FormatString = "Sr. No.|CASE NO.       | REQUISITION NO.      |TRANSFER TIME| STATUS  "
'                                End If
'                                GoTo MOVE:
'                            End If
'                        End If
'                    End If
'                Else
'                    MSFlexGrid1.TextMatrix(I, 4) = "SAVE ERROR"
'                    MYDELAY (5)
'                    If MSFlexGrid1.Rows <> 2 Then
'                        MSFlexGrid1.RemoveItem I
'                    Else
'                        MSFlexGrid1.Clear
'                        MSFlexGrid1.FormatString = "Sr. No.|CASE NO.       | REQUISITION NO.      |TRANSFER TIME| STATUS  "
'                    End If
'                    GoTo MOVE:
'                End If
'
'                If MSFlexGrid1.Rows <> 2 Then
'                    MSFlexGrid1.RemoveItem I
'                Else
'                    MSFlexGrid1.Clear
'                    MSFlexGrid1.FormatString = "Sr. No.|CASE NO.       | REQUISITION NO.      |TRANSFER TIME| STATUS   "
'                End If
'MOVE:       Next I2
'
'            If MSFlexGrid1.Rows = 2 Then
'                If ERROR = True Or FILLERROR = True Or SAVEERROR = True Then
'                    TXT_MESSAGE.Text = "THERE IS AN ERROR WITH SOME OF THE RECORDS PLEASE CLICK ON THE ERROR LOG PRINT BUTTON"
'                    CMD_ERRLOGPRINT.Enabled = True
'                    CMD_EXIT.Enabled = False
'                Else
'                    CMD_ERRLOGPRINT.Enabled = False
'                    TXT_MESSAGE.Text = "  NOW YOU CAN QUIT PLEASE CLICK ON EXIT BUTTON!!!!!"
'                    MYDELAY (5)
'                End If
'            End If
'retry:
'        If ENDOFPROCESS = True Then
'            ROUTINE = True
'        Else
'            MYDELAY (10)
'            FILLGRID
'        End If
'
'    Loop
'
'    Unload Me
'    End

End Sub



Private Sub Form_Load()
'MonthView1.Visible = False
'MonthView2.Visible = False
 ProgressBar1.Visible = False
If App.PrevInstance = True Then End

srno = 1
'LBL_PHASE.Caption = PH
'MsgBox PH
' Set OBJ = New HMSCrypt
On Error GoTo MSG:

'    Dim OBJ As Object
'
'    For Each OBJ In CBCCOMMIT.Controls
'        If TypeOf OBJ Is CommandButton Then SetWindowRgn OBJ.hWnd, CreateRoundRectRgn(2, 2, 80, 23, 50, 20), True
'    Next

    DATETIME

    ipadd = Winsock1.LocalIP

    msk_sysdate.Text = Format(DATE11, "DD/MM/YYYY")
    msk_systime.Text = Format(Time, "HH:MM:SS")

 txt_frm.Text = Format(DATE11, "DD/MM/YYYY")
  txt_to.Text = Format(DATE11, "DD/MM/YYYY")

 'for dept selection
    str = "select dept_name, dept_code from PATOOLSLIB.ma106"
    Set sel_deptqr = rdb.CreateQuery("micro_select2", str)

    'FOR MODALITY SELECTION
    str = "select modlitydes,modalitycd from rislib.rmodality WHERE DEPT_CODE=? "
    Set sel_moadlityqr = rdb.CreateQuery("micro_SELECT3", str)

str = "SELECT PRO_CODE FROM DISLIB.DMULPROCE WHERE REQNO=?"
    Set SELDMUL_QR = rdb.CreateQuery("", str)


Set sel_deptrs = sel_deptqr.OpenResultset(1, 3)
    While Not sel_deptrs.EOF
        cmb_dept.AddItem Trim(sel_deptrs("dept_name")) & "$" & Trim(sel_deptrs("dept_code"))
        sel_deptrs.MoveNext
    Wend
    For I = 0 To (cmb_dept.ListCount - 1)
        If Trim(cmb_dept.List(I)) = "MICROBIOLOGY$GZZ" Then
                cmb_dept.ListIndex = I
        End If
    Next I
     If Trim(cmb_dept) = "" Then
        Exit Sub
    End If
    'copied and pasted
    I = InStr(cmb_dept.Text, "$")
    txt_deptcode = Mid(cmb_dept.Text, I + 1, 3)
    sel_moadlityqr(0) = Trim(txt_deptcode)
    Set sel_modalityrs = sel_moadlityqr.OpenResultset(1, 3)
    If sel_modalityrs.RowCount > 0 Then
        cmb_modality.clear
        While Not sel_modalityrs.EOF
            cmb_modality.AddItem Trim(sel_modalityrs("modlitydes")) & "$" & Trim(sel_modalityrs("modalitycd"))
            sel_modalityrs.MoveNext
        Wend
        sel_modalityrs.Close
    Else
        sel_modalityrs.Close
    End If



    TXT_COMMENTS1.AddItem " "
    TXT_COMMENTS1.AddItem "THIS HbsAg  RATE NOTE IS CLOSE TO THE CUT OFF VALUE,CUT OFF : 0.05  PATIENT VALUE:       IU/mL. PLEASE REPEAT THE SAMPLE FOR HBsAg TESTING. "
    TXT_COMMENTS1.AddItem "THIS HCV RATE NOTE IS CLOSE TO THE CUT OFF VALUE,CUT OFF : 1.00 PATIENT VALUE:        . PLEASE REPEAT THE SAMPLE FOR HCV PCR."
    TXT_COMMENTS1.AddItem "PLEASE REPEAT THE SAMPLE FOR HIV TESTING AFTER 4 WEEKS.          "
    TXT_COMMENTS1.AddItem "The HBsAg value is low and requires a confirmation test. Please send a fresh sample for HBV DNA Quantitive."
    TXT_COMMENTS1.AddItem "Please send a Whole blood sample for Quantitive PCR."
    TXT_COMMENTS1.AddItem "Please send EDATA sample for Quantitative PCR."
    TXT_COMMENTS1.AddItem "Please send EDATA sample for Qualitative PCR."
    TXT_COMMENTS1.AddItem "SPECIMENS WITH CONCENTRATION VALUES <6.0 AU/mL ARE CONSIDERED NONREACTIVE FOR IgG ANTIBODIES TO CMV."
    TXT_COMMENTS1.AddItem "SPECIMENS WITH CONCENTRATION VALUES >=6.0 AU/mL ARE CONSIDERED REACTIVE FOR IgG ANTIBODIES TO CMV. THIS INDICATES PAST OR ACUTE INFECTION."
    TXT_COMMENTS1.AddItem "HBsAg TEST HAS BEEN PERFORMED AT NICHOLAS PIRAMAL LABORATORIES." 'Shraddha 03/08/10



    TXT_COMMENTS.AddItem " "
    TXT_COMMENTS.AddItem "THIS HbsAg  RATE NOTE IS CLOSE TO THE CUT OFF VALUE,CUT OFF : 0.05  PATIENT VALUE:       IU/mL. PLEASE REPEAT THE SAMPLE FOR HBsAg TESTING. "
    TXT_COMMENTS.AddItem "THIS HCV RATE NOTE IS CLOSE TO THE CUT OFF VALUE,CUT OFF : 1.00 PATIENT VALUE:        . PLEASE REPEAT THE SAMPLE FOR HCV PCR."
    TXT_COMMENTS.AddItem "PLEASE REPEAT THE SAMPLE FOR HIV TESTING AFTER 4 WEEKS.          "
    TXT_COMMENTS.AddItem "The HBsAg value is low and requires a confirmation test. Please send a fresh sample for HBV DNA Quantitive."
    TXT_COMMENTS.AddItem "Please send a Whole blood sample for Quantitive PCR."
    TXT_COMMENTS.AddItem "Please send EDATA sample for Quantitative PCR."
    TXT_COMMENTS.AddItem "Please send EDATA sample for Qualitative PCR."
    TXT_COMMENTS.AddItem "SPECIMENS WITH CONCENTRATION VALUES <6.0 AU/mL ARE CONSIDERED NONREACTIVE FOR IgG ANTIBODIES TO CMV."
    TXT_COMMENTS.AddItem "SPECIMENS WITH CONCENTRATION VALUES >=6.0 AU/mL ARE CONSIDERED REACTIVE FOR IgG ANTIBODIES TO CMV. THIS INDICATES PAST OR ACUTE INFECTION."
    TXT_COMMENTS.AddItem "HBsAg TEST HAS BEEN PERFORMED AT NICHOLAS PIRAMAL LABORATORIES." 'Shraddha 03/08/10


    str = "insert into dislib.microproce values(?,?,?,?,?,?,?,?,?,?)"
    Set ins_micro = rdb.CreateQuery("ins", str)

    str = "insert into dislib.MICROTEMP values(?,?,?,?,?,?,?,?,?,?)"
    Set ins_micro_temp = rdb.CreateQuery("ins_temp", str)

    str = "delete from dislib.MICROTEMP"
    Set dlt_micro = rdb.CreateQuery("dlt", str)


    str = "select DISTINCT(REQNO),CASENO from INTERFACE.Minterface where iflag='2' AND TRIM(MACHINENO)='V5600'  and trim(fld2)='NonReactive'"
    Set selhintfqr = rdb.CreateQuery("selhintfqr", str)

    str = "select DISTINCT(REQNO),CASENO from INTERFACE.Minterface where iflag='2' AND TRIM(MACHINENO)='V5600'  and trim(fld2)='NonReactive'"
    Set selhintfqr1 = rdb.CreateQuery("selectrec", str)
    'SHWETA
    'STR = "select * from INTERFACE.Minterface where reqno=? and iflag='2' AND TRIM(MACHINENO)='V5600' AND trim(fld2)='NonReactive' ORDER BY TTIME"
    str = "select * from INTERFACE.Minterface where reqno=? and iflag='2' ORDER BY TTIME"
    Set qr_minterface = rdb.CreateQuery("qr_m", str)

    str = "SELECT * FROM INTERFACE.HINTERFACE WHERE REQNO=?"
    Set SELHINTFVAL_QR = rdb.CreateQuery("SELHINTFVAL_QR", str)

    str = "INSERT INTO INTERFACE.CBCINTFERR (REQNO,ERRORCODE,ERRORMODE,ERRORDESC,ERRORDATE,ERRORTIME,PRINTFLAG) VALUES(?,?,?,?,?,?,?)"
    Set INSCBCINTFERR_QR = rdb.CreateQuery("INSCBCINTFERR_QR", str)

    str = "SELECT PATNSTAT FROM RISLIB.RREQUISITN WHERE REQNNO=?"
    Set SELPATNSTAT_QR = rdb.CreateQuery("SELPATNSTAT_QR", str)

    str = "SELECT PRVDIAGDES FROM PATOOLSLIB.TA1012 WHERE CASE_NO=?"
    Set selta1012_qr = rdb.CreateQuery("SELTA1012_QR", str)

    str = "SELECT * FROM INTERFACE.TIMESLOT"
    Set SELTIMESLOT_QR = rdb.CreateQuery("SELTIMESLOT_QR", str)

    str = "INSERT INTO DISLIB.DHEMCBCDET (REQNO,PRO_CODE,OBSERVALUE,OBSER_TYPE,SOURCE,ENTRYBY,REMARKS,UOM,NRANGE,PVALUE) VALUES(?,?,?,?,?,?,?,?,?,?)"
    Set INSDHEMCBCDET_QR = rdb.CreateQuery("INSDHEMCBCDET_QR", str)

    str = "SELECT * FROM DISLIB.DHEMCBCRNG WHERE PROCODE=? AND AGE2=? AND SEX=? "
    Set SELDHEMCBCQR = rdb.CreateQuery("SELDHEMCBCQR", str)

    str = "SELECT * FROM INTERFACE.CBCINTFERR WHERE ERRORDATE=? AND PRINTFLAG=?"
    Set SELCBCINTFERR_QR = rdb.CreateQuery("SELCBCINTFERR_QR", str)

    str = "UPDATE INTERFACE.CBCINTFERR SET PRINTFLAG=? WHERE REQNO=? AND ERRORCODE=? AND ERRORMODE=? AND ERRORDATE=? AND ERRORTIME=?"
    Set UPDCBCINTFERR_QR = rdb.CreateQuery("UPDCBCINTFERR_QR", str)

    str = "update INTERFACE.Minterface set iflag = '3' where reqno = ? AND TESTRESULT<>''"
    Set updhinterface_qr = rdb.CreateQuery("updhinter", str)

'SHWETA 11-09-2012
    str = "update INTERFACE.Minterface set TESTRESULT = ? ,FLD2=? where reqno = ? AND TESTCODE=?"
    Set UPD_MINTERFACE_qr = rdb.CreateQuery("MINTERFACE_qr", str)


    str = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
    Set updrrequisitn1_qr = rdb.CreateQuery("updrrequisitn1_qr1NHE", str)

    str = "SELECT CASENO FROM RISLIB.RREQUISITN WHERE REQNNO=?"
    Set SELRREQUISITN_QR = rdb.CreateQuery("SELRREQUISITN_QR", str)

    str = "UPDATE RISLIB.RTIMELOG SET COMREPDATE=?,COMREPTIME=?,PRNREPDATE=?,PRNREPTIME=? WHERE REQNNO=?"
    Set UPDRTIMEC_QR = rdb.CreateQuery("QWEQWEQWEccy", str)

    str = "SELECT CASENO,PATNSTAT FROM RISLIB.RREQUISITN WHERE REQNNO=?"
    Set SELCASENO_QR = rdb.CreateQuery("SELCASENO_QR", str)

    str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=? "
    Set SELSTAFF_QR = rdb.CreateQuery("SELSTAFF_QR", str)

    str = "SELECT * FROM PATOOLSLIB.TA101 WHERE CASE_NO=?"
    Set SELSEXAGE_QR = rdb.CreateQuery("SELSEXAGE_QR", str)

    '***** Shraddha 27/09/10 *****
    str = "SELECT * FROM DISLIB.MICROPROCE WHERE REQNO=? and PRO_CODE=?"
    Set SEL_MICROPROCE_QR = rdb.CreateQuery("", str)

    str = "SELECT * FROM DISLIB.MICROTEMP WHERE REQNO=? and PRO_CODE=?"
    Set SEL_MICROTEMP_QR = rdb.CreateQuery("", str)
    '*****************************
    str = "INSERT INTO dislib.DMBCCOMM (REQNO,COMMENTS) VALUES (?,?)"
    Set ins_dcbcrep1_qr = rdb.CreateQuery("", str)

     str = "SELECT * FROM DISLIB.DMBCCOMM WHERE REQNO=?"
    Set SEL_DMBCCOMM_QR = rdb.CreateQuery("", str)

    str = "UPDATE DISLIB.DMBCCOMM SET COMMENTS=? WHERE REQNO=?"
    Set UPD_DMBCCOMM_qr = rdb.CreateQuery("", str)

    str = "INSERT INTO DISLIB.MICRO_LIST VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
    Set INSMICROLIST_QR = rdb.CreateQuery("", str)

    str = "SELECT * FROM DISLIB.MICRO_LIST WHERE REQNO=?"
    Set SELMICROLIST_QR = rdb.CreateQuery("", str)


    str = "DELETE FROM DISLIB.MICRO_LIST WHERE REQNO=?"
    Set DELMICROLIST_QR = rdb.CreateQuery("", str)

    str = "UPDATE DISLIB.MICROPROCE SET OBSERVALUE=?,RESULT=?,NRANGE=?,FLD2=?,FLD1=? WHERE REQNO=? AND PRO_CODE=?"
    Set updmproce_qr = rdb.CreateQuery("", str)

    str = "UPDATE DISLIB.MICROTEMP SET OBSERVALUE=?,RESULT=? ,FLD1=? WHERE REQNO=? AND PRO_CODE=?"
    Set updmtemp_qr = rdb.CreateQuery("", str)


    '''''charu 08082012


    'SECURITY STARTS

    str = "SELECT * FROM DISLIB.DISLNKUSR WHERE USR_ID = ? "
    Set USRNAME_QR = rdb.CreateQuery("USRNAME_QR", str)

       ' USRNAME_QR(0) = "A" + logonfrm.usrid
       USRNAME_QR(0) = logonfrm.usrid
        Set USRNAME_RS = USRNAME_QR.OpenResultset(1, 1)

        If USRNAME_RS.RowCount > 0 Then
            LOGINUSRNAME = Trim(USRNAME_RS("USERNAME"))
            LOGINTYP = Trim(USRNAME_RS("USRTYPE"))
            LOGINUSRID = Trim(USRNAME_RS("usr_id"))
            TXT_OPERATOR.Text = (LOGINUSRNAME)
            TXT_OPERATOR.Enabled = False
        Else
            MsgBox "The Security Details of this User is not available!!", vbInformation, "DIS"
            Exit Sub
        End If

    'SECURITY ENDS

    Set SELTIMESLOT_RS = SELTIMESLOT_QR.OpenResultset(1, 1)
    TXT_MINUTES.Text = Trim(SELTIMESLOT_RS("TIMESLOT"))

    Exit Sub

MSG:
    'MsgBox Err.Number
    'MsgBox Err.Description

End Sub






Sub DISPLAYINTXT(ByVal row As Integer)
Exit Sub
txt_srno.Text = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 0))
txt_pathno.Text = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))
TXT_CASENO.Text = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2))
txt_reqno.Text = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3))
txt_vhbsag.Text = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4))
txt_RHBsAg.Text = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5))
txt_VHCV.Text = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6))
txt_RHCV.Text = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
txt_VHIV.Text = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 8))
txt_RHIV.Text = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 9))
'TXT_DCOMM.Text = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.ROW, 10))

lbl_name.Caption = GETNAME(TXT_CASENO.Text)
 'If Trim(TXT_COMMENTS.Text) <> "" Then
    SEL_DMBCCOMM_QR(0) = "GZZSE" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3))
    Set SEL_DMBCCOMM_RS = SEL_DMBCCOMM_QR.OpenResultset(2, 1)
    If Not SEL_DMBCCOMM_RS.EOF Then
    'Else
    On Error Resume Next

    TXT_DCOMM.Text = Trim(SEL_DMBCCOMM_RS(1))

    End If
'End If
COLORTEXTBOX txt_RHBsAg
COLORTEXTBOX txt_RHCV
COLORTEXTBOX txt_RHIV

End Sub
Sub COLORTEXTBOX(t1 As TextBox)
If Trim(UCase(t1.Text)) = "REACTIVE" Or Trim(UCase(t1.Text)) = "HIGH REACTIVE" Then
t1.ForeColor = vbRed
t1.FontBold = True
End If

If Trim(UCase(t1.Text)) = "NONREACTIVE" Then
t1.ForeColor = &H8000&
t1.FontBold = True
End If


If Trim(UCase(t1.Text)) = "GRAYZONE" Then
t1.ForeColor = vbBlack
t1.FontBold = True
End If

If Trim(UCase(t1.Text)) = "BORDERLINE" Then
t1.ForeColor = vbBlue
t1.FontBold = True
End If

End Sub










Public Sub MSFlexGrid1_Click()
case_no_tst = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2))
req_no_tst = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
If MSFlexGrid1.col = 1 Then
mod_reqno = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
MSFlexGrid1.col = 6
MSFlexGrid1.CellBackColor = &HC0FFC0
MSFlexGrid1.col = 2
MSFlexGrid1.CellBackColor = &HC0FFC0
MSFlexGrid1.col = 3
MSFlexGrid1.CellBackColor = &HC0FFC0
MSFlexGrid1.col = 4
MSFlexGrid1.CellBackColor = &HC0FFC0
MSFlexGrid1.col = 5
MSFlexGrid1.CellBackColor = &HC0FFC0
MSFlexGrid1.col = 1
MSFlexGrid1.CellBackColor = &HC0FFC0
PRINTING
End If

If MSFlexGrid1.col = 6 Then
   If Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6)) = "No" Then
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6) = "Yes"
  MSFlexGrid1.CellBackColor = &H80FF&
  I = MSFlexGrid1.row
MSFlexGrid1.row = I
UPDATE
flag = "URG"
'INSERTTOLISTTBL
'End If 'COMM BY SHWETA 06-03-2013 AS PER MICRO DEPT. SUGGESTION
flag = "YES"
Else
mod_reqno = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
 str = "UPDATE RISLIB.RREQUISITN SET REQNSTAT=?,COMMITFLAG=? WHERE REQNNO=?"
    Set UPDRREQ_QR = rdb.CreateQuery("", str)
    UPDRREQ_QR(0) = "T"
UPDRREQ_QR(1) = Null
UPDRREQ_QR(2) = mod_reqno
UPDRREQ_QR.Execute
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6) = "No"
 MSFlexGrid1.CellBackColor = &HFFFF&
  I = MSFlexGrid1.row
MSFlexGrid1.row = I
End If

End If

If MSFlexGrid1.col = 5 Then
        If Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6)) = "Yes" Then
        MsgBox ("Report is already commited.")
Exit Sub
Else

cbccommit.Hide
Print MSFlexGrid1.CellForeColor = &HC0FFC0
micro_no = MSFlexGrid1.row

'ADDED BY SHWETA 23/09/2021
Dim RdmulproceQR As rdoQuery
Dim RdmulproceRS As rdoResultset
REQNO = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
 str = "SELECT * FROM DISLIB.dmulproce WHERE reqno = ? and pro_code in (select ser_code from dislib.tmtestcode where type='J')"
    Set RdmulproceQR = rdb.CreateQuery("", str)

        RdmulproceQR(0) = REQNO
        Set RdmulproceRS = RdmulproceQR.OpenResultset(1, 1)

        If RdmulproceRS.RowCount > 0 Then
          MSFlexGrid1.CellBackColor = &H8000000D
        'Load (Report_microbiology)
       frm_mo.show

      frm_mo.cmd_view_Click
       frm_mo.CMD_NEW.Enabled = False
         frm_mo.cmd_view.Enabled = False
         If IsNumeric(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)) Then
          frm_mo.opt_staff.Value = True
          frm_mo.Opt_patient.Value = False
            frm_mo.MaskEdBox1.Mask = "#######"
         Else
            frm_mo.opt_staff.Value = False
          frm_mo.Opt_patient.Value = True
            frm_mo.MaskEdBox1.Mask = "??>/#####"
         End If
       'frm_mo.txt_spm.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4)
          frm_mo.MaskEdBox1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
          frm_mo.cmb_reqno.AddItem MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7)

            LEN1 = Len(Trim(cbccommit.MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7)))

                    If LEN1 = 13 Then
                    res_reqno = Left(cbccommit.MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7), 3) & "/" & Mid(cbccommit.MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7), 4, 2) & "/" & Mid(cbccommit.MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7), 6, 2) & "/" & Right(cbccommit.MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7), 6)
                     End If
                    
                    If LEN1 = 16 Then
                    res_reqno = Left(cbccommit.MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7), 3) & Mid(cbccommit.MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7), 5, 2) & Mid(cbccommit.MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7), 8, 2) & Right(cbccommit.MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7), 6)
                    End If
                    
                    If LEN1 = 27 Then
                    res_reqno = Left(cbccommit.MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7), 3) & Mid(cbccommit.MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7), 5, 2) & Mid(cbccommit.MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7), 8, 2) & Mid(cbccommit.MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7), 11, 6)
                    End If
                  '  frm_mo.cmb_reqno.AddItem(res_reqno & "$" & MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
        frm_mo.cmb_reqno.Text = Replace(res_reqno, "/", "") '& "$" & MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
         '  frm_mo.cmb_modality.Text = cmb_modality.Text
         '    frm_mo.Text2.Text = res_reqno & "$" & MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
        frm_mo.cmb_reqno_Click
Exit Sub
               End If


'''




  MSFlexGrid1.CellBackColor = &H8000000D
        'Load (Report_microbiology)
       Report_microbiology1.show

      Report_microbiology1.cmd_view_Click
       Report_microbiology1.CMD_NEW.Enabled = False
         Report_microbiology1.cmd_view.Enabled = False
         If IsNumeric(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)) Then
          Report_microbiology1.opt_staff.Value = True
          Report_microbiology1.Opt_patient.Value = False
            'Report_microbiology1.MaskEdBox1.Mask = "#######"
         Else
            Report_microbiology1.opt_staff.Value = False
          Report_microbiology1.Opt_patient.Value = True
          '  Report_microbiology1.MaskEdBox1.Mask = "??>/#####"
         End If
       Report_microbiology1.txt_spm.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4)
        Report_microbiology1.MaskEdBox1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
          Report_microbiology1.cmb_reqno.AddItem MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7)

      LEN1 = Len(Trim(cbccommit.MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7)))

If LEN1 = 13 Then
res_reqno = Left(cbccommit.MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7), 3) & "/" & Mid(cbccommit.MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7), 4, 2) & "/" & Mid(cbccommit.MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7), 6, 2) & "/" & Right(cbccommit.MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7), 6)
 End If

If LEN1 = 16 Then
res_reqno = Left(cbccommit.MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7), 3) & Mid(cbccommit.MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7), 5, 2) & Mid(cbccommit.MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7), 8, 2) & Right(cbccommit.MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7), 6)
End If

If LEN1 = 27 Then
res_reqno = Left(cbccommit.MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7), 3) & Mid(cbccommit.MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7), 5, 2) & Mid(cbccommit.MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7), 8, 2) & Mid(cbccommit.MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7), 11, 6)
End If
        Report_microbiology1.cmb_reqno.Text = res_reqno & "$" & MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
           Report_microbiology1.cmb_modality.Text = cmb_modality.Text
             Report_microbiology1.Text2.Text = res_reqno & "$" & MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
        Report_microbiology1.cmb_reqno_Click

               End If

 End If

End Sub

Private Sub MSFlexGrid1_Scroll()
TXT_COMMENTS.Visible = False
'CHK_Save.Visible = False
End Sub

Private Sub Timer1_Timer()

    DATETIME
    msk_sysdate.Text = Format(DATE11, "DD/MM/YYYY")
    msk_systime.Text = Format(time1, "HH:MM:SS")

End Sub

Public Sub Main()

Dim er As rdoError

On Error GoTo CnEh
Set renv = rdoEnvironments(0)
renv.CursorDriver = rdUseOdbc

    Set rdb = renv.OpenConnection(DSNAME:="DIS_MUL", Prompt:=rdDriverNoPrompt, ReadOnly:=False, Connect:="System=10.80.1.2;UID=PATOOLS;PWD=PATOOLS;Database = EC_TMH")

Exit Sub

CnEh:
        Debug.Print Err, ERROR

        For Each er In rdoErrors
            Debug.Print er.Description, er.Number
        Next er
            MsgBox "Either this user dosen't exist on Server or " & Chr(13) & " there is a communication link failure"
            MsgBox (Err)
        End


End Sub

Public Sub DATETIME()

    Set rs1 = rdb.OpenResultset("SELECT DISTINCT CURRENT DATE DT1 , CURRENT TIME TM1 FROM PATOOLSLIB.MA107")
    DATE11 = rs1("DT1")
    DATE11 = Format(DATE11, "DD/MM/YYYY")
    time1 = rs1("TM1")

End Sub

Public Sub UPDATE()

'On Error GoTo MSG:

    Dim rno As String
'CD/23095


'If CBCCOMMIT.LBL_PHASE.Caption = "SAVE" Then
''    str = "UPDATE RISLIB.RREQUISITN SET  REQNSTAT='T',PREPORTDT=CURRENT DATE,PREPORTBY=? WHERE REQNNO=?"
''    Set updrrequisitnt_qr = rdb.CreateQuery("", str)
''    updrrequisitnt_qr(0) = Trim(TXT_OPERATOR.Text)
''    updrrequisitnt_qr(1) = "GZZSE" & Trim(Trim(MSFlexGrid1.TextMatrix(i, 3)))
''    updrrequisitnt_qr.Execute
''    UPDATEMINTERFACE
'Else

    rno = Trim(Trim(MSFlexGrid1.TextMatrix(I, 7)))
    'Set IFLAG 3
    updhinterface_qr(0) = rno  'selhintfrs("REQNO")
    updhinterface_qr.Execute


    updrrequisitn1_qr(0) = Trim(TXT_OPERATOR.Text)
    updrrequisitn1_qr(1) = Trim(msk_sysdate.FormattedText)
    updrrequisitn1_qr(2) = "F"
    updrrequisitn1_qr(3) = "Y"
    updrrequisitn1_qr(4) = Trim(Trim(MSFlexGrid1.TextMatrix(I, 7)))  '"UZZHL" + selhintfrs("REQNO")
    updrrequisitn1_qr.Execute

    UPDRTIMEC_QR(0) = Trim(msk_sysdate.FormattedText)
    UPDRTIMEC_QR(1) = Trim(msk_systime.FormattedText)
    UPDRTIMEC_QR(2) = Trim(msk_sysdate.FormattedText)
    UPDRTIMEC_QR(3) = Trim(msk_systime.FormattedText)
    UPDRTIMEC_QR(4) = Trim(Trim(MSFlexGrid1.TextMatrix(I, 7))) '"UZZHL" + selhintfrs("REQNO")
    UPDRTIMEC_QR.Execute
    'UPDATEMINTERFACE
'End If
    Exit Sub

MSG:

    INSCBCINTFERR_QR(0) = req_chglen(Trim(MSFlexGrid1.TextMatrix(I, 7)))
    INSCBCINTFERR_QR(1) = Err.Number
    INSCBCINTFERR_QR(2) = "COMMIT"
    INSCBCINTFERR_QR(3) = Err.Description
    INSCBCINTFERR_QR(4) = Date
    INSCBCINTFERR_QR(5) = Time
    INSCBCINTFERR_QR(6) = "N"
  '  INSCBCINTFERR_QR.Execute
    ERROR = True

End Sub
Sub UPDATEMINTERFACE()
UPD_MINTERFACE_qr(0) = MSFlexGrid1.TextMatrix(I, 4)
UPD_MINTERFACE_qr(1) = MSFlexGrid1.TextMatrix(I, 5)
UPD_MINTERFACE_qr(2) = "GZZSE" & Trim(Trim(MSFlexGrid1.TextMatrix(I, 3)))
UPD_MINTERFACE_qr(3) = "G129"
UPD_MINTERFACE_qr.Execute


UPD_MINTERFACE_qr(0) = MSFlexGrid1.TextMatrix(I, 6)
UPD_MINTERFACE_qr(1) = MSFlexGrid1.TextMatrix(I, 7)
UPD_MINTERFACE_qr(2) = "GZZSE" & Trim(Trim(MSFlexGrid1.TextMatrix(I, 3)))
UPD_MINTERFACE_qr(3) = "G134"
UPD_MINTERFACE_qr.Execute

UPD_MINTERFACE_qr(0) = MSFlexGrid1.TextMatrix(I, 8)
UPD_MINTERFACE_qr(1) = MSFlexGrid1.TextMatrix(I, 9)
UPD_MINTERFACE_qr(2) = "GZZSE" & Trim(Trim(MSFlexGrid1.TextMatrix(I, 3)))
UPD_MINTERFACE_qr(3) = "G171"
UPD_MINTERFACE_qr.Execute

End Sub


Public Sub MYDELAY(inttime As String)



Dim sttime
Dim endtime1 As Long
Dim TESTTIME As Long


'sttime = Timer

endtime1 = Timer + inttime
TESTTIME = Timer

Do While Timer < endtime1
 '   Text1 = endtime1
 '   Text2 = TESTTIME
    DoEvents
  '  TESTTIME = Timer
Loop


End Sub

Public Sub PRINTING()

Dim VAR1 As String
If Len(Trim(mod_reqno)) > 0 Then
    VAR1 = Trim(Replace(mod_reqno, "/", ""))    'MaskEdBox1.FormattedText
    Dim strLink As String
    'strLink = "http://intranet.tmc.gov.in/emrweb/default.asp" '10.1.4.9
  '  strLink = "http://intranet.tmc.gov.in/EMRREPORTS/allreports.aspx?reqno=" & VAR1 & "~guest"
    'MRUNAL
    strLink = "http://intranet.hbchrcm.tmc.gov.in/emrreports/allreports.aspx?reqno=" & VAR1 & "~print"
    'END
    Shell "c:\program files\internet explorer\iexplore " & strLink
Else
    MsgBox "Please Enter The Case No."
End If

Exit Sub



    'On Error GoTo MSG:

    Dim str1 As String
    Dim str2 As String
'Dim REN As String
'REN = "GZZSE" & RTrim(MSFlexGrid1.TextMatrix(I21, 3)) & "'"
  '  DoEvents

Dim procode As String

    SELDMUL_QR(0) = mod_reqno
    Set SELDMUL_RS = SELDMUL_QR.OpenResultset(1, 1)
    If SELDMUL_RS.EOF = False Then
    procode = Trim(SELDMUL_RS(0))
    End If

Dim SELREP_QR As rdoQuery
Dim SELREP_RS As rdoResultset
Dim REPORT As String
    str = "SELECT REPORT FROM DISLIB.DMOPRISMER WHERE REQNO=?"
    Set SELREP_QR = rdb.CreateQuery("", str)
    SELREP_QR(0) = mod_reqno
    Set SELREP_RS = SELREP_QR.OpenResultset(1, 1)
    If SELREP_RS.EOF = False Then
        If IsNull(SELREP_RS(0)) = False Then

    REPORT = Trim(SELREP_RS(0))
    End If

    End If
PatInfo MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7), mod_reqno, ""


pno = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))
'wbval = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5))
If UCase(Trim(wbval)) = "OUTPATIENT" Then
wbval = ""
End If
For I1 = 0 To 20 'Increased from 10 to 20 on 05/08/09 by Shraddha Pandit
    CRYSTALREPORT1.Formulas(I1) = ""
Next I1

For I1 = 0 To 20 'Increased from 10 to 20 on 05/08/09 by Shraddha Pandit
    CrystalReport2.Formulas(I1) = ""
Next I1


If txt_modcode = "SE" Then
    If procode = "G261" Or procode = "G262" Or procode = "G267" Or procode = "G264" Then
    'OTHER_REPORT
    RPT_Other_report
    ElseIf procode = "G129" Or procode = "G134" Or procode = "G171" Or procode = "G126" Then
     'SEROLOGY
    Else
    RPT_Other_report

    End If


ElseIf txt_modcode = "BA" Then


    If procode = "G204" Or procode = "G206" Or procode = "G209" Or procode = "G256" Or procode = "G265" Or procode = "G266" Or procode = "G268" Or procode = "G111" Then
    'OTHER_REPORT
    RPT_bacteriology_con
    ElseIf procode = "G115" Or Len(Trim(REPORT)) > 1 Then
    'PATHO_NO_REPORT
    RPT_patho_no
    Else
    'BACTERIOLOGY_REPORT
     rpt_bacteriology_report
    End If



ElseIf txt_modcode = "MB" Then

 If procode = "G253" Then
 'OTHER REPORT
 RPT_Other_report
 Else
 'mruanl_3_8_15
  If procode = "G108" Then
RPT_bacteriology_con
Else
 'MYCO
 RPT_myco
End If
'eND
  End If


ElseIf txt_modcode = "CM" Then
  If procode = "G256" Or procode = "G268" Then
    'BACTERIOLOGY_CON
    RPT_bacteriology_con
    ElseIf procode = "G101" Then
    'URINE_REPORT
     RPT_urine
    Else
    RPT_stool
    'STOOL REPORT
   End If



ElseIf txt_modcode = "MC" Then

    If procode = "G151" Then
    'MICRO_MYCO
    RPT_MICRO_MYCO
    Else
    'OTHER
     RPT_Other_report
   End If
Else
  'OTHER
  RPT_Other_report
End If






'    If IsNumeric(MSFlexGrid1.TextMatrix(I21, 2)) Then
'    'for staff
'    str1 = IPADD & " " & Format(date11, "DD/MM/YYYY") & " " & Time
'    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & "GZZSE" & Trim(MSFlexGrid1.TextMatrix(I21, 3)) & "'"
'    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\SEROLOGYstaff.rpt"
'    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
'    CrystalReport1.Formulas(0) = "REQNO = '" & MSFlexGrid1.TextMatrix(I21, 2) & "'"
'    CrystalReport1.Formulas(12) = "USRID = '" & str1 & "'"
'
'    Else
'    'for patient
'    str1 = IPADD & " " & Format(date11, "DD/MM/YYYY") & " " & Time
'    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & "GZZSE" & Trim(MSFlexGrid1.TextMatrix(I21, 3)) & "'"
'    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\SEROLOGY.rpt"
'    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
'    CrystalReport1.Formulas(0) = "REQNO = '" & "GZZSE" & MSFlexGrid1.TextMatrix(I21, 3) & "'"
'    CrystalReport1.Formulas(12) = "USRID = '" & str1 & "'"
'    End If
'
'    'CrystalReport1.Destination = crptToPrinter
'    CrystalReport1.WindowState = crptMaximized
'    CrystalReport1.Destination = crptToWindow
'
'    'CrystalReport1.RecordsPrinted
'    'CrystalReport1.RecordsRead
'    CrystalReport1.Action = 1
'
   ' dlt_micro.Execute
    Exit Sub

MSG:

'    INSCBCINTFERR_QR(0) = "GZZSE" & req_chglen(Trim(MSFlexGrid1.TextMatrix(I, 3)))
'    INSCBCINTFERR_QR(1) = Err.Number
'    INSCBCINTFERR_QR(2) = "PRINT"
'    INSCBCINTFERR_QR(3) = Err.Description
'    INSCBCINTFERR_QR(4) = Date
'    INSCBCINTFERR_QR(5) = Time
'    INSCBCINTFERR_QR(6) = "N"
'    INSCBCINTFERR_QR.Execute
'    ERROR = True

'    Resume Next

End Sub
Sub rpt_bacteriology_report()
CRYSTALREPORT1.SelectionFormula = "{bacterolf.REQNNO} = '" & mod_reqno & "'"
                    'SHWETA 06-SEPT-2010
                    CRYSTALREPORT1.ReportFileName = "c:\DISREPORTSECLF\bacteriology_report.rpt"
                    CRYSTALREPORT1.Formulas(1) = "proFIN = 'MICROBIOLOGY REPORT'"
                    CRYSTALREPORT1.Formulas(2) = "ENT = '" & "Entered By   " & Trim(cmb_enteryby) & "'"
                    CRYSTALREPORT1.Formulas(3) = "USRID = '" & "" & "'"
                    CRYSTALREPORT1.Formulas(4) = "patho = '" & "Yes" & "'"
                    CRYSTALREPORT1.Formulas(8) = "nam = '" & MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) & "'"
                    CRYSTALREPORT1.Formulas(9) = "STDOB = '" & sac & "'"
                    CRYSTALREPORT1.Formulas(7) = "WARD = '" & Trim(wbval) & "'"
                    'CRYSTALREPORT1.Formulas(5) = "ICU = '" & "Yes" & "'"
                     'SELECT FOR ORGANISM ISOLATED

                    '***** Shraddha Pandit 30/06/09 *****
'                    If Len(Trim(txt_organ1.Text)) > 0 And Len(Trim(TXT_ORGAN2.Text)) = 0 And Len(Trim(TXT_ORGAN3.Text)) = 0 Then
'                        CrystalReport1.Formulas(5) = "ORGAN1 = '" & "Yes" & "'"
'                    ElseIf Len(Trim(txt_organ1.Text)) > 0 And Len(Trim(TXT_ORGAN2.Text)) > 0 And Len(Trim(TXT_ORGAN3.Text)) = 0 Then
'                        CrystalReport1.Formulas(5) = "ORGAN2 = '" & "Yes" & "'"
'                    ElseIf Len(Trim(txt_organ1.Text)) > 0 And Len(Trim(TXT_ORGAN2.Text)) > 0 And Len(Trim(TXT_ORGAN3.Text)) > 0 Then
'                        CrystalReport1.Formulas(5) = "ORGAN3 = '" & "Yes" & "'"
''                    End If
'CrystalReport1.Formulas(5) = ""

Dim SELbact_QR As rdoQuery
Dim SELLbact_RS As rdoResultset
Dim REPORT As String
    str = "SELECT * FROM DISLIB.bacterolf WHERE REQNO=?"
    Set SELbact_QR = rdb.CreateQuery("", str)
    SELbact_QR(0) = mod_reqno
    Set SELLbact_RS = SELbact_QR.OpenResultset(1, 1)
    If SELLbact_RS.EOF = False Then

                    If Len(Trim(SELLbact_RS("orgso1"))) > 0 And (Len(Trim(SELLbact_RS("orgso2"))) = 0 Or IsNull(SELLbact_RS("orgso2")) = True) And (Len(Trim(SELLbact_RS("orgso3"))) = 0 Or IsNull(SELLbact_RS("orgso3")) = True) Then
                        CRYSTALREPORT1.Formulas(5) = "ORGAN1 = '" & "Yes" & "'"
                    ElseIf Len(Trim(SELLbact_RS("orgso1"))) > 0 And Len(Trim(SELLbact_RS("orgso2"))) > 0 And (Len(Trim(SELLbact_RS("orgso3"))) = 0 Or IsNull(SELLbact_RS("orgso3")) = True) Then
                        CRYSTALREPORT1.Formulas(5) = "ORGAN2 = '" & "Yes" & "'"
                    ElseIf Len(Trim(SELLbact_RS("orgso1"))) > 0 And Len(Trim(SELLbact_RS("orgso2"))) > 0 And Len(Trim(SELLbact_RS("orgso3"))) > 0 Then
                        CRYSTALREPORT1.Formulas(5) = "ORGAN3 = '" & "Yes" & "'"
                    End If
    End If
                    '************************************
                    Set TEMP_RS = rdb.OpenResultset("SELECT * FROM DISLIB.DMULPROCE WHERE  REQNO = '" & mod_reqno & "'")
                    If TEMP_RS.EOF = False Then
                        If Trim(UCase(TEMP_RS("PRO_CODE"))) = "G204" Or Trim(UCase(TEMP_RS("PRO_CODE"))) = "G205" Or Trim(UCase(TEMP_RS("PRO_CODE"))) = "G206" Then
                        CRYSTALREPORT1.Formulas(6) = "TXTPRIMARY_STAIN = ''"
                        Else
                        CRYSTALREPORT1.Formulas(6) = "TXTPRIMARY_STAIN = 'Primary Smear (Gram Stain) '"
                        End If
                    End If
                      CRYSTALREPORT1.WindowState = crptMaximized
                    CRYSTALREPORT1.Connect = "ODBC;Database=DISLIB;UID=tmh;PWD= tmh;DSN=DISDSN"
                    CRYSTALREPORT1.Action = 1
End Sub
Sub RPT_myco()
CRYSTALREPORT1.SelectionFormula = "{MYCOLF.REQNNO} = '" & mod_reqno & "'"
                    CRYSTALREPORT1.ReportFileName = "c:\DISREPORTSECLF\Myco.rpt"
                    ''''
                    CRYSTALREPORT1.Formulas(1) = "proFIN = 'MICROBIOLOGY REPORT'"
                    CRYSTALREPORT1.Formulas(2) = "ENT = '" & "Entered By   " & Trim(cmb_enteryby) & "'"
                    CRYSTALREPORT1.Formulas(3) = "USRID = '" & str & "'"
                    CRYSTALREPORT1.Formulas(13) = "BEDNO ='" & Trim(wbval) & "'"
                    CRYSTALREPORT1.Formulas(14) = "REPORT1='" & REPORTSTR & "'"
                    CRYSTALREPORT1.Formulas(4) = "nam='" & MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) & "'"
                    CRYSTALREPORT1.Formulas(5) = "STDOB='" & sac & "'" '
                    CRYSTALREPORT1.Formulas(6) = "PATHNUMBER='" & Trim(pno) & "'" '
                    CRYSTALREPORT1.WindowState = crptMaximized
                    CRYSTALREPORT1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                    CRYSTALREPORT1.Action = 1
End Sub
Sub RPT_stool()
CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & mod_reqno & "'"
                    CRYSTALREPORT1.ReportFileName = "c:\DISREPORTSECLF\Stool_Report.rpt"
                    CRYSTALREPORT1.Formulas(1) = "proFIN = 'MICROBIOLOGY REPORT'"
                    CRYSTALREPORT1.Formulas(2) = "CODE2 = 'Stool Routine Examination '"
                    CRYSTALREPORT1.Formulas(3) = "ENT = '" & "Entered By   " & Trim(cmb_enteryby) & "'"
                    CRYSTALREPORT1.Formulas(4) = "USRID = '" & str & "'"
                    CRYSTALREPORT1.Formulas(13) = "WARD='" & Trim(wbval) & "'"
                    CRYSTALREPORT1.Formulas(14) = "nam='" & MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) & "'"
                    CRYSTALREPORT1.Formulas(15) = "STDOB='" & sac & "'" '
                    CRYSTALREPORT1.Formulas(16) = "PATHNUMBER='" & Trim(pno) & "'" '
                    'If SEL_DMUL_RS("PRO_CODE") = "G251" Then
                    'Commented and added by Shraddha on 04/07/2011


                    If Trim(SELDMUL_RS("PRO_CODE")) = "G251" Or Trim(SELDMUL_RS("PRO_CODE")) = "G102" Then
                    Dim comment As String
'                    comment = ""
'                    comment = "Stool samples collected during menstrual bleeding, constipation induced bleeding, bleeding hemorrhoids or when rectal medication is used may cause positive results" & vbCrLf
'                    comment = comment & "Certain medications may induce gastroinstestinal bleeding and cause false positive reactions hence should be avoided during and prior to the testing period" & vbCrLf
'                    comment = comment & "Diet containing exogenous peroxidases may induce false positive results" & vbCrLf
'                    comment = comment & "Dosages of Vitamin C more than 250mg per day will cause false negative results"

                        'CrystalReport1.Formulas(17) = "OccultComment='" & comment & "'"
                        CRYSTALREPORT1.Formulas(17) = "OccultComment='Yes'"
                    Else
                        CRYSTALREPORT1.Formulas(17) = "OccultComment=''"
                    End If

                      CRYSTALREPORT1.WindowState = crptMaximized
                    CRYSTALREPORT1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                    CRYSTALREPORT1.Action = 1

End Sub
Sub RPT_urine()
CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & mod_reqno & "'"
                    CRYSTALREPORT1.ReportFileName = "c:\DISREPORTSECLF\Urine_Report.rpt"
                    CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & mod_reqno & "'"
                    CRYSTALREPORT1.Formulas(0) = "REQNO = '" & Trim(mod_reqno) & "'"
                    CRYSTALREPORT1.Formulas(1) = "proFIN = 'MICROBIOLOGY REPORT'"
                    CRYSTALREPORT1.Formulas(2) = "CODE1 = 'Urine Routine Examination '"
                    CRYSTALREPORT1.Formulas(3) = "ENT = '" & "Entered By   " & Trim(cmb_enteryby) & "'"
                    CRYSTALREPORT1.Formulas(4) = "USRID = '" & str & "'"
                    CRYSTALREPORT1.Formulas(7) = "pathnumber='" & Trim(pno) & "'"
                    CRYSTALREPORT1.Formulas(8) = "ward='" & wbval & "'"     ' txt_ward & " /" & txt_bedno & "'"
                    CRYSTALREPORT1.Formulas(9) = "NAM='" & MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) & "'"
                    CRYSTALREPORT1.Formulas(10) = "STDOB='" & sac & "'" 'txt_sex & " /" & txt_age & "'"
                    'If Opt_patient.Value Then
                   '     CrystalReport1.Formulas(11) = "ST='" & txt_category & " /" & txt_status & "'"
                  '  Else
                  '      CrystalReport1.Formulas(11) = "ST='Staff /" & txt_status & "'"
                   ' End If
                    CRYSTALREPORT1.Formulas(13) = "bedno='" & txt_bedno & "'"
                    CRYSTALREPORT1.WindowState = crptMaximized
                    CRYSTALREPORT1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                    CRYSTALREPORT1.Action = 1
                    '****************************
'                    CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & mod_reqno & "'"
'                    CRYSTALREPORT1.ReportFileName = "c:\DISREPORTSEC\Stool_Report.rpt"
'                    CRYSTALREPORT1.Formulas(1) = "proFIN = 'MICROBIOLOGY REPORT'"
'                    CRYSTALREPORT1.Formulas(2) = "CODE2 = 'Stool Routine Examination '"
'                    CRYSTALREPORT1.Formulas(3) = "ENT = '" & "Entered By   " & Trim(cmb_enteryby.Text) & "'"
'                    CRYSTALREPORT1.Formulas(4) = "USRID = '" & str & "'"
'                    CRYSTALREPORT1.Formulas(13) = "bedno='" & txt_bedno & "'"
End Sub
Sub RPT_patho_no()
                    CRYSTALREPORT1.SelectionFormula = "{PATHONOLF.REQNO} = '" & mod_reqno & "'"
                    CRYSTALREPORT1.ReportFileName = "c:\DISREPORTSECLF\patho_no_report.rpt"
                    CRYSTALREPORT1.Formulas(1) = "proFIN = 'MICROBIOLOGY REPORT'"
                    CRYSTALREPORT1.Formulas(2) = "ENT = '" & "Entered By   " & Trim(cmb_enteryby) & "'"
                    CRYSTALREPORT1.Formulas(3) = "USRID = '" & str & "'"
                    CRYSTALREPORT1.Formulas(5) = "nam = '" & MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) & "'"
                    CRYSTALREPORT1.Formulas(6) = "dob= '" & sac & "'"
                    CRYSTALREPORT1.Formulas(7) = "ward= '" & wbval & "'"
                    CRYSTALREPORT1.Formulas(8) = "pathnumber= '" & Trim(pno) & "'"
                    '*********ADDED BY SHWETA 13-AUG-09
                    'Dim TEMP_RS As rdoResultset
'                    SELSTR = "SELECT * FROM DISLIB.DMULPROCE WHERE  REQNO = '" & mod_reqno & "'"
                    Set TEMP_RS = rdb.OpenResultset("SELECT * FROM DISLIB.DMULPROCE WHERE  REQNO = '" & mod_reqno & "'")
                    If TEMP_RS.EOF = False Then
                        If Trim(UCase(TEMP_RS("PRO_CODE"))) = "G204" Or Trim(UCase(TEMP_RS("PRO_CODE"))) = "G205" Or Trim(UCase(TEMP_RS("PRO_CODE"))) = "G206" Then
                        CRYSTALREPORT1.Formulas(4) = "TXTPRIMARY_STAIN = ''"
                        Else
                        CRYSTALREPORT1.Formulas(4) = "TXTPRIMARY_STAIN = 'Primary Smear (Gram Stain) '"
                        End If
                    End If
                    TEMP_RS.Close
                    CRYSTALREPORT1.WindowShowPrintBtn = True
        CRYSTALREPORT1.WindowShowRefreshBtn = True
        CRYSTALREPORT1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CRYSTALREPORT1.WindowState = crptMaximized
        CRYSTALREPORT1.Action = 1

End Sub
Sub RPT_MICRO_MYCO()
 CrystalReport2.ReportFileName = "C:\DISREPORTSECLF\MICRO_MYCO.rpt"
        CrystalReport2.WindowTitle = "Microbiology Report"

        CrystalReport2.SelectionFormula = "{drequisitn.REQNO} = '" & mod_reqno & "'"
      '  CrystalReport2.Formulas(0) = "name ='" & MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) & "'"
        'CrystalReport2.Formulas(1) = "sex/age ='" & Trim(sac) & "'"

       ' CrystalReport2.Formulas(3) = "USRID = '" & UI & "'"
       ' CrystalReport2.Formulas(4) = "comment = '" & Trim(txt_MCcomment) & "'"
        CrystalReport2.Destination = crptToWindow 'SHWETA 13-AUG-09
        'CrystalReport2.Destination = crptToPrinter 'SHRADDHA 09-11-09
        'CrystalReport2.Destination = crptToPrinter
        CrystalReport2.WindowShowPrintBtn = True
        CrystalReport2.WindowShowRefreshBtn = True
        CrystalReport2.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport2.WindowState = crptMaximized
        CrystalReport2.Action = 1
End Sub
Sub RPT_bacteriology_con()
CrystalReport2.SelectionFormula = "{RREQUISITN.REQNNO} = '" & mod_reqno & "'"
        CrystalReport2.ReportFileName = "c:\DISREPORTSECLF\bacteriology_con.rpt"
        CrystalReport2.Formulas(0) = "proFIN = 'MICROBIOLOGY REPORT'"
        CrystalReport2.Formulas(1) = "ENT = '" & "Entered By   " & Trim(logonfrm.txtusername.Text) & "'"
        CrystalReport2.Formulas(2) = "patho = '" & "Yes" & "'"
        CrystalReport2.Formulas(3) = "bedno='" & txt_bedno & "'"
        CrystalReport2.Formulas(4) = "REQNO = '" & Trim(cmb_reqno) & "'"
        CrystalReport2.Formulas(5) = "pathnumber='" & Trim(pno) & "'"
        CrystalReport2.Formulas(6) = "NAM='" & MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) & "'"
        CrystalReport2.Formulas(7) = "STDOB='" & sac & "'" ' txt_sex & " /" & txt_age & "'"
        CrystalReport2.Formulas(8) = "ward='" & Trim(wbval) & "'" ' & " / " & txt_status & " '"
        CrystalReport2.Formulas(9) = "USRID = '" & UI & "'"
        Set TEMP_RS = rdb.OpenResultset("SELECT * FROM DISLIB.DMULPROCE WHERE  REQNO = '" & mod_reqno & "'")
        If TEMP_RS.EOF = False Then
            If Trim(UCase(TEMP_RS("PRO_CODE"))) = "G204" Or Trim(UCase(TEMP_RS("PRO_CODE"))) = "G205" Or Trim(UCase(TEMP_RS("PRO_CODE"))) = "G206" Then
            CRYSTALREPORT1.Formulas(10) = "TXTPRIMARY_STAIN = ''"
            Else
            CRYSTALREPORT1.Formulas(10) = "TXTPRIMARY_STAIN = 'Primary Smear (Gram Stain) '"
            End If
            If TEMP_RS("PRO_CODE") = "G256" Then
            CrystalReport2.Formulas(10) = "PComment='Yes'"
            Else
                CrystalReport2.Formulas(10) = "PComment=''"
            End If
        End If

        CrystalReport2.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport2.WindowState = crptMaximized
        CrystalReport2.Action = 1

End Sub
Sub RPT_Other_report()
 CrystalReport2.SelectionFormula = "{REPORTOTLF.REQNO} = '" & mod_reqno & "'"
        'CrystalReport2.ReportFileName = "c:\DISREPORTSEClf\GZZSER_REPORT.RPT"
        CrystalReport2.ReportFileName = "c:\DISREPORTSEClf\other_report.rpt"
        CrystalReport2.Formulas(0) = "proFIN = 'MICROBIOLOGY REPORT'"
        CrystalReport2.Formulas(1) = "ENT = '" & "Entered By   " '& Trim(cmb_enteryby.Text) & "'"
        CrystalReport2.Formulas(3) = "bedno='" & txt_bedno & "'"
        CrystalReport2.Formulas(4) = "REQNO = '" & Trim(mod_reqno) & "'"
        CrystalReport2.Formulas(5) = "pathnumber='" & Trim(pno) & "'"
        CrystalReport2.Formulas(6) = "NAM='" & MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) & "'"
        CrystalReport2.Formulas(7) = "STDOB='" & Trim(sac) & "'" 'txt_sex & " /" & txt_age & "'"
       'CrystalReport2.Formulas(8) = "ST='c/ip'" '& txt_category & " /" & txt_status & "'"
        CrystalReport2.Formulas(9) = "USRID = '" & UI & "'"
        Set TEMP_RS = rdb.OpenResultset("select * from dislib.gzzserolog WHERE  REQNO = '" & mod_reqno & "'", 1, 1)
        If Not TEMP_RS.EOF Then
            If Trim(TEMP_RS("Result")) = "" Then CrystalReport2.Formulas(10) = "flgresult='false'"
            If Trim(TEMP_RS("INTERPRETN")) = "" Then CrystalReport2.Formulas(11) = "flginter='false'"
            If Trim(TEMP_RS("METHOD")) = "" Then CrystalReport2.Formulas(12) = "flgMethod='false'"
            If Trim(TEMP_RS("LIMITATION")) = "" Then CrystalReport2.Formulas(13) = "flgLimitation='false'"
        End If


                CrystalReport2.Formulas(14) = "ward='" & Trim(wbval) & "'"   'txt_ward & " /" & txt_bedno & "'"


        CrystalReport2.Connect = "ODBC;Database=DISLIB;UID=tmh;PWD= tmh;DSN=DISDSN"
        CrystalReport2.WindowState = crptMaximized
        CrystalReport2.Action = 1
End Sub
Public Sub FILLGRID()
INSGRID = True
'If CBCCOMMIT.LBL_PHASE.Caption = "COMMIT" Then
MSFlexGrid1.FormatString = "SNo.| Micro No.                       |   Case No.     |   Patient Name                     | Specimen                  | Edit    | Committed |REQ Noo                                                               "
' CMD_MODIFY.Enabled = True
 Command3.Enabled = True
 'Else
 'MSFlexGrid1.FormatString = "SNo.|Micro No.|Case No.  | Req No.   | Result      | REPORT           |  Saved   |Urgent "
 'CMD_MODIFY.Enabled = False
 Command3.Enabled = True 'AS TOLD BY KALPANA MAM
' End If
Dim RESULTRS1 As rdoResultset
Dim sql1 As String
'On Error GoTo fillgriderr:
Dim t1 As String, T2 As String, t3 As String
Dim tt1 As String, tt2 As String, tt3 As String, pathno As String
''''''''''    FILLING FLEXGRID
Dim GZZSPSERO_QR As rdoQuery
Dim GZZSPSERO_RS As rdoResultset
    MSFlexGrid1.Rows = 1
    I = 1
    TXT_MESSAGE.Text = "  RETRIEVING DETAILS PLEASE WAIT......"
    'MYDELAY (5)

'    Set selhintfrs = rdb.OpenResultset("select DISTINCT(REQNO),CASENO from INTERFACE.Minterface where iflag='2' AND TRIM(MACHINENO)='V5600'  and trim(fld2)='NonReactive'", 1, 3)
'SEPT-2012 SHWETA

I = InStr(cmb_modality.Text, "$")
txt_modcode = Mid(cmb_modality.Text, I + 1, 2)

Dim selhintf_qr As rdoQuery
Dim selhintfrs As rdoResultset
Dim GZZREQPATH_QR As rdoQuery
Dim GZZREQPATH_RS As rdoResultset
RR = msk_sysdate.Text
FDT = RR


If txt_frm.Text = "" Then
Set selhintf_qr = rdb.CreateQuery("", "select * FROM  DISLIB.RREQDREQLF WHERE DEPT_CODE ='GZZ' AND MODALITYCD='" & Trim(txt_modcode) & "' AND REQNSTAT='T' AND (DAYS(CURRENT DATE) - DAYS(PREPORTDT)) <= 07 ORDER BY PATHNO")
Else
Set selhintf_qr = rdb.CreateQuery("", "select * FROM  DISLIB.RREQDREQLF WHERE DEPT_CODE ='GZZ' AND MODALITYCD='" & Trim(txt_modcode) & "' AND REQNSTAT='T' AND (PREPORTDT>= '" & Format(txt_frm.Text, "YYYY-MM-DD") & "' and PREPORTDT <= '" & Format(txt_to.Text, "YYYY-MM-DD") & "') ORDER BY PATHNO")
End If

    If click = "2" Then
    If txt_frm.Text = "" Then
Set selhintf_qr = rdb.CreateQuery("", "select * FROM  DISLIB.RREQDREQLF WHERE DEPT_CODE ='GZZ' AND MODALITYCD='" & Trim(txt_modcode) & "' AND REQNSTAT='F' AND (DAYS(CURRENT DATE) - DAYS(FREPORTDT)) <= 07 ORDER BY PATHNO")
Else
Set selhintf_qr = rdb.CreateQuery("", "select * FROM  DISLIB.RREQDREQLF WHERE DEPT_CODE ='GZZ' AND MODALITYCD='" & Trim(txt_modcode) & "' AND REQNSTAT='F' AND (FREPORTDT>= '" & Format(txt_frm.Text, "YYYY-MM-DD") & "' and FREPORTDT <= '" & Format(txt_to.Text, "YYYY-MM-DD") & "') ORDER BY PATHNO")
End If
    End If

    'selhintf_qr(0) = "GZZ"
    'selhintf_qr(1) = Trim(txt_modcode)
    Set selhintfrs = selhintf_qr.OpenResultset(1, 1)


    str = "SELECT PATHNO,icds2 FROM DISLIB.DREQUISITN WHERE REQNO=?"
    Set GZZREQPATH_QR = rdb.CreateQuery("", str)


     srno = 1
    If selhintfrs.RowCount <> 0 Then
        selhintfrs.MoveFirst

        VAR_REQNO = selhintfrs("REQNNO")

     Dim CNT As Integer
     Dim a As Integer

     CNT = selhintfrs.RowCount


FILL:       While selhintfrs.EOF = False
        resultflag = False


         VAR_REQNO = selhintfrs("REQNNO")
        '''
         GZZREQPATH_QR(0) = VAR_REQNO
         Set GZZREQPATH_RS = GZZREQPATH_QR.OpenResultset(1, 1)
         pathno = Trim(GZZREQPATH_RS("PATHNO"))
         patstat = Trim(selhintfrs("PATNSTAT"))
         PatInfo selhintfrs("CASENO"), VAR_REQNO, patstat
          If Trim(GZZREQPATH_RS("icds2")) = "" Or IsNull(GZZREQPATH_RS("icds2")) = True Then
          ICU = ""
          Else
                If GZZREQPATH_RS("icds2") = "CG" Then
                ICU = "Casualty General"
                ElseIf GZZREQPATH_RS("icds2") = "CM" Then
                ICU = "Casualty Main"
                ElseIf GZZREQPATH_RS("icds2") = "I1" Then
                ICU = "ICU 1st floor"
                ElseIf GZZREQPATH_RS("icds2") = "RMC" Then
                ICU = "RMC"
                ElseIf GZZREQPATH_RS("icds2") = "-" Then
                ICU = "-"
                Else
                ICU = "ICU"
                End If

          End If
         C1 = selhintfrs("CASENO")
          wb = ""
            If patstat = "IP" Then
                   Dim selwb_qr As rdoQuery
                   Dim selwb_rs As rdoResultset
                   str = "SELECT * FROM patoolslib.TA105 WHERE CASE_NO=? "
                   Set selwb_qr = rdb.CreateQuery("", str)
                   selwb_qr(0) = selhintfrs("CASENO")

                   Set selwb_rs = selwb_qr.OpenResultset(1, 1)
                   If selwb_rs.EOF = False Then
                   wb = selwb_rs("WARD_NAME") & " / " & Trim(selwb_rs("bed_no"))
                   Else
                   If Trim(ICU) = "" Then
                   wb = "Out Patient"
                   End If
                   End If
            End If
            SELDMUL_QR(0) = VAR_REQNO
            Set SELDMUL_RS = SELDMUL_QR.OpenResultset(1, 1)
            If SELDMUL_RS.EOF = False Then
            procode = Trim(SELDMUL_RS(0))
            End If




''FOLLOWING CODE COMMENTED BY SHWETA 14-02-2013
'    If procode = "G129" Or procode = "G134" Or procode = "G132" Or procode = "G171" Or procode = "G126" Or procode = "G130" Or procode = "G131" Or procode = "G133" Or procode = "G136" Then
'    Else
'          MSFlexGrid1.AddItem srno & vbTab & Trim(pathno) & vbTab & selhintfrs("CASENO") & vbTab & Mid(Trim(VAR_REQNO), 1) & vbTab & PATDETAILS & vbTab & icu & wb & vbTab & " View Report" & vbTab & TR
'          srno = srno + 1
'    End If
    str = "SELECT TESTCODE FROM DISLIB.SPSEROPF WHERE TESTCODE=?"
    Set GZZSPSERO_QR = rdb.CreateQuery("", str)
        GZZSPSERO_QR(0) = procode
          Set GZZSPSERO_RS = GZZSPSERO_QR.OpenResultset(1, 1)
          If GZZSPSERO_RS.EOF Then
            Dim specimn
       str = "SELECT specimen FROM dislib.dmsponanti WHERE reqno='" & Mid(Trim(VAR_REQNO), 1) & "'"
    Set SELCBCINTFERR_QR_sp = rdb.CreateQuery("", str)
    Set SELCBCINTFERR_RS_sp = SELCBCINTFERR_QR_sp.OpenResultset(1, 1)
    If SELCBCINTFERR_RS_sp.RowCount > 0 Then
 specimen = SELCBCINTFERR_RS_sp(0)
               End If

             If click = "2" Then
                 MSFlexGrid1.AddItem srno & vbTab & Trim(pathno) & vbTab & selhintfrs("CASENO") & vbTab & PATDETAILS & vbTab & specimen & vbTab & "Edit" & vbTab & " Yes " & vbTab & Mid(Trim(VAR_REQNO), 1)
             Else
               MSFlexGrid1.AddItem srno & vbTab & Trim(pathno) & vbTab & selhintfrs("CASENO") & vbTab & PATDETAILS & vbTab & specimen & vbTab & "Edit" & vbTab & " No " & vbTab & Mid(Trim(VAR_REQNO), 1)
             End If
            cmb_srno.AddItem (Mid(Trim(VAR_REQNO), 1) & "$" & Trim(pathno) & "$" & specimen)
                  srno = srno + 1
            End If
NXTROW:
                    selhintfrs.MoveNext
        Wend

        TXT_MESSAGE.Text = "  NO. OF RECORDS FOUND ARE......" & srno ' - 1 '& selhintfrs.RowCount
        'MYDELAY (6)

    Else
        TXT_MESSAGE.Text = "  NO RECORDS FOUND FOR COMMITTING .....!!!!!!"
        'MYDELAY (2)
        TXT_MESSAGE.Text = "  NOW YOU CAN QUIT PLEASE CLICK ON EXIT BUTTON!!!!!"

    End If
    If MSFlexGrid1.Rows > 1 Then
    DISPLAYINTXT (1)
    End If
    Exit Sub

fillgriderr:
            INSCBCINTFERR_QR(0) = selhintfrs("REQNO")
            INSCBCINTFERR_QR(1) = I
            INSCBCINTFERR_QR(2) = "FILLERROR"
            INSCBCINTFERR_QR(3) = "CASE NUMBER NOT FOUND FOR THIS REQ. NO."
            INSCBCINTFERR_QR(4) = Date
            INSCBCINTFERR_QR(5) = Time
            INSCBCINTFERR_QR(6) = "N"
            INSCBCINTFERR_QR.Execute
            FILLERROR = True
            selhintfrs.MoveNext
            GoTo FILL:
'UPTO HERE
End Sub
Public Sub micro()



qr_minterface(0) = "GZZSE" & Trim(Trim(MSFlexGrid1.TextMatrix(I, 3)))
Set rs_minterface = qr_minterface.OpenResultset(1, 1)
Do While rs_minterface.EOF = False

    '***** Check where rec already exists : Shraddha 27/09/10 *****
    SEL_MICROPROCE_QR(0) = "GZZSE" & Trim(Trim(MSFlexGrid1.TextMatrix(I, 3)))
    SEL_MICROPROCE_QR(1) = Trim(rs_minterface(3))
    Set SEL_MICROPROCE_RS = SEL_MICROPROCE_QR.OpenResultset(1, 1)

    If SEL_MICROPROCE_RS.EOF Then
    '**************************************************************

        key1 = makekey
        ins_micro(0) = "GZZSE" & Trim(Trim(MSFlexGrid1.TextMatrix(I, 3)))
        ins_micro(1) = rs_minterface(3)
        ins_micro(2) = obj.Encrypt(rs_minterface(6), key)
        ins_micro(3) = obj.Encrypt(rs_minterface(15), key)
        ins_micro(4) = TXT_OPERATOR
        ins_micro(5) = ""
        ins_micro(6) = rs_minterface(15) 'shweta to save decrypted data 03-08-2012
        ins_micro(7) = key1
        ins_micro(8) = rs_minterface(6) 'shweta to save decrypted data 03-08-2012
        ins_micro(9) = ""
        ins_micro.Execute
    Else

    End If


    'SEL_MICROTEMP_QR
    '***** Check where rec already exists : Shraddha 27/09/10 *****
    SEL_MICROTEMP_QR(0) = "GZZSE" & Trim(Trim(MSFlexGrid1.TextMatrix(I, 3)))
    SEL_MICROTEMP_QR(1) = Trim(rs_minterface(3))
    Set SEL_MICROTEMP_RS = SEL_MICROTEMP_QR.OpenResultset(1, 1)

    If SEL_MICROTEMP_RS.EOF Then
    '**************************************************************


    ins_micro_temp(0) = "GZZSE" & Trim(Trim(MSFlexGrid1.TextMatrix(I, 3)))
    ins_micro_temp(1) = rs_minterface(3)
    ins_micro_temp(2) = rs_minterface(6)
    ins_micro_temp(3) = rs_minterface(15)
    ins_micro_temp(4) = TXT_OPERATOR
    ins_micro_temp(5) = ""
    ins_micro_temp(6) = ""
    ins_micro_temp(7) = key1
    ins_micro_temp(8) = ""
    ins_micro_temp(9) = ""
    ins_micro_temp.Execute


    End If

    rs_minterface.MoveNext
Loop
End Sub
Public Function makekey()
Dim Y, m, d, hr, mn, sc As Integer
Dim myarr(1 To 20) As String
Dim STQ As String
key = 0

d = Val(Mid(DATE11, 1, 2))
m = Val(Mid(DATE11, 4, 2))
Y = Val(Mid(DATE11, 9, 2))
hr = Val(Mid(time1, 1, 2))
mn = Val(Mid(time1, 4, 2))
sc = Val(Mid(time1, 7, 2))

Dim MyValue
MyValue = Int((50 * rnd) + 1)
key = Y + m + d + hr + 2 * sc - mn + MyValue
Select Case key
    Case 1 To 50
        key = key + 25
    Case 1 To 100
        key = key - 30
    Case 101 To 150
        key = key - 60
    Case 151 To 200
        key = key - 80
    Case 201 To 250
        key = key - 130
    Case 251 To 253
        key = key - 140
    Case Is > 253
        key = 90
 End Select
 makekey = key

 End Function


Public Sub micro_temp()

End Sub

Private Sub TXT_COMMENTS_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
        If Trim(TXT_COMMENTS.Text) <> "" Then
             SEL_DMBCCOMM_QR(0) = "GZZSE" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3))
            Set SEL_DMBCCOMM_RS = SEL_DMBCCOMM_QR.OpenResultset(2, 1)
            If Not SEL_DMBCCOMM_RS.EOF Then
                UPD_DMBCCOMM_qr(0) = Trim(TXT_COMMENTS.Text)
                UPD_DMBCCOMM_qr(1) = "GZZSE" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3))
                UPD_DMBCCOMM_qr.Execute
            Else
            ins_dcbcrep1_qr(0) = "GZZSE" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3))
            ins_dcbcrep1_qr(1) = Trim(cbo_comments.Text)
           ' ins_dcbcrep1_qr.Execute
            End If

            MSFlexGrid1.TextMatrix((MSFlexGrid1.row), 10) = Trim(TXT_COMMENTS.Text)
        End If
        TXT_COMMENTS.Visible = False
    End If
End Sub



Private Sub TXT_COMMENTS1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

  If Trim(TXT_COMMENTS1.Text) <> "" Then
             SEL_DMBCCOMM_QR(0) = "GZZSE" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3))
            Set SEL_DMBCCOMM_RS = SEL_DMBCCOMM_QR.OpenResultset(2, 1)
            If Not SEL_DMBCCOMM_RS.EOF Then
                UPD_DMBCCOMM_qr(0) = Trim(TXT_COMMENTS1.Text)
                UPD_DMBCCOMM_qr(1) = "GZZSE" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3))
                UPD_DMBCCOMM_qr.Execute
            Else
            ins_dcbcrep1_qr(0) = "GZZSE" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3))
            ins_dcbcrep1_qr(1) = Trim(TXT_COMMENTS1.Text)
            ins_dcbcrep1_qr.Execute
        End If
            TXT_DCOMM.Text = Trim(TXT_COMMENTS1.Text)
        End If

        TXT_COMMENTS1.Visible = False

    End If
End Sub

Private Sub TXT_DATA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Trim(TXT_DATA.Text) <> "" Then

            MSFlexGrid1.TextMatrix((MSFlexGrid1.row), MSFlexGrid1.col) = Trim(TXT_DATA.Text)
        End If
        TXT_DATA.Visible = False
    End If
End Sub

Private Sub TXT_DCOMM_GotFocus()
TXT_COMMENTS1.Visible = True
End Sub


Private Sub txt_frm_Change()
'MonthView1.Visible = True
'MonthView2.Visible = True
End Sub

Private Sub txt_RHBsAg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txt_RHCV.SetFocus
End If
End Sub


Private Sub txt_RHCV_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txt_RHIV.SetFocus
End If
End Sub

Private Sub txt_RHIV_Change()
If KeyAscii = 13 Then
TXT_DCOMM.SetFocus
End If
End Sub

