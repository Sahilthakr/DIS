VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form Em_lab 
   Caption         =   "Emergency"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13980
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8745
   ScaleWidth      =   13980
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chk_bloodgas 
      Caption         =   "Blood Gas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   67
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   480
      TabIndex        =   52
      Top             =   4800
      Width           =   10815
      Begin VB.CheckBox chk_livfuntest 
         Caption         =   "Liver Function Tests"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   0
         TabIndex        =   62
         Top             =   0
         Width           =   2535
      End
      Begin VB.CheckBox chk_prot 
         Caption         =   "Proteins"
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
         Height          =   255
         Left            =   720
         TabIndex        =   61
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox chk_albumin 
         Caption         =   "Albumin"
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
         Height          =   255
         Left            =   2880
         TabIndex        =   60
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox chk_glob 
         Caption         =   "Globulins"
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
         Height          =   255
         Left            =   5400
         TabIndex        =   59
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox chk_alkphos 
         Caption         =   "Alk.Phosphatase"
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
         Height          =   255
         Left            =   7200
         TabIndex        =   58
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox chk_totbil 
         Caption         =   "Total Bilirubin"
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
         Height          =   255
         Left            =   9120
         TabIndex        =   57
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox chk_dirbil 
         Caption         =   "Direct Bilirubin"
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
         Height          =   255
         Left            =   720
         TabIndex        =   56
         Top             =   840
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox Chk_SGOT1 
         Caption         =   "S.G.O.T"
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
         Height          =   195
         Left            =   5400
         TabIndex        =   55
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox chk_sgpt 
         Caption         =   "S.G.P.T"
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
         Height          =   255
         Left            =   7200
         TabIndex        =   54
         Top             =   840
         Width           =   1095
      End
      Begin VB.CheckBox chk_indirbil 
         Caption         =   "Indirect Bilirubin"
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
         Height          =   255
         Left            =   2880
         TabIndex        =   53
         Top             =   840
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Haematological Investigations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1695
      Left            =   480
      TabIndex        =   38
      Top             =   6240
      Visible         =   0   'False
      Width           =   10815
      Begin VB.CheckBox chk_pcv 
         Caption         =   "PCV"
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
         Left            =   8520
         TabIndex        =   49
         Top             =   720
         Width           =   975
      End
      Begin VB.Frame Frame9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   8400
         TabIndex        =   44
         Top             =   360
         Width           =   1575
         Begin VB.CheckBox chk_btct 
            Caption         =   "BT - CT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   0
            TabIndex        =   48
            Top             =   0
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox CHK_CT 
            Caption         =   "CT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   720
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CheckBox chk_bt 
            Caption         =   "BT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   45
            Top             =   360
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.Frame Frame8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         TabIndex        =   42
         Top             =   360
         Width           =   4575
         Begin VB.CheckBox chk_pc 
            Caption         =   "Platelet Count"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2640
            TabIndex        =   51
            Top             =   360
            Width           =   1815
         End
         Begin VB.CheckBox Chk_wbc 
            Caption         =   "W.B.C. Count"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2640
            TabIndex        =   50
            Top             =   840
            Width           =   1575
         End
         Begin VB.CheckBox CHK_hbpcv 
            Caption         =   "CBC"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   0
            TabIndex        =   47
            Top             =   0
            Width           =   1575
         End
         Begin VB.CheckBox chk_hb 
            Caption         =   "Hb"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   43
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.CheckBox CHK_tt 
         Caption         =   "Thrombin Time"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   41
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox CHK_PTT 
         Caption         =   "PTTK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7200
         TabIndex        =   40
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox CHK_PT 
         Caption         =   "PT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   39
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Blood Gas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1935
      Left            =   9360
      TabIndex        =   35
      Top             =   2880
      Visible         =   0   'False
      Width           =   1935
      Begin VB.CheckBox chk_bgel 
         Caption         =   "Electrolytes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CheckBox chk_venous 
         Caption         =   "Venous"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   1200
         Width           =   975
      End
      Begin VB.CheckBox chk_cap 
         Caption         =   "Capillary"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   840
         Width           =   1095
      End
      Begin VB.CheckBox chk_am 
         Caption         =   "Arterial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.ComboBox cmb_nmeterial 
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
      ItemData        =   "Em_lab.frx":0000
      Left            =   2880
      List            =   "Em_lab.frx":0002
      TabIndex        =   33
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9840
      TabIndex        =   28
      Top             =   1560
      Width           =   3975
      Begin VB.CheckBox chk_ur24hr 
         Caption         =   "Urine Osmolality (24 Hours)"
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
         Height          =   375
         Left            =   2160
         TabIndex        =   71
         Top             =   840
         Width           =   1935
      End
      Begin VB.CheckBox chk_urrandom 
         Caption         =   "Urine Osmolality (Random)"
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
         Height          =   375
         Left            =   2160
         TabIndex        =   70
         Top             =   240
         Width           =   2295
      End
      Begin VB.CheckBox chk_urea 
         Caption         =   "Urea"
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
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox chk_creat 
         Caption         =   "Creatinine"
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
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmd_back 
      Caption         =   "<<&Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      TabIndex        =   24
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3720
      TabIndex        =   19
      Top             =   3360
      Width           =   5535
      Begin VB.CheckBox chk_phosphorus 
         Caption         =   "Phosphorus"
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
         Height          =   255
         Left            =   1920
         TabIndex        =   73
         Top             =   960
         Width           =   1455
      End
      Begin VB.CheckBox chk_lipase 
         Caption         =   "Lipase"
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
         Height          =   255
         Left            =   240
         TabIndex        =   72
         Top             =   960
         Width           =   975
      End
      Begin VB.CheckBox chk_uricacid 
         Caption         =   "Uric Acid"
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
         Height          =   255
         Left            =   3840
         TabIndex        =   69
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox chk_ca1 
         Caption         =   "Calcium"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox chk_ca 
         Caption         =   "Ionic Calcium"
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
         Height          =   255
         Left            =   1920
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox chk_amylase 
         Caption         =   "Amylase"
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
         Height          =   375
         Left            =   1920
         TabIndex        =   22
         Top             =   600
         Width           =   1095
      End
      Begin VB.CheckBox chk_lactate 
         Caption         =   "Lactate"
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
         Height          =   375
         Left            =   3840
         TabIndex        =   21
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox chk_magn 
         Caption         =   "Magnesium"
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
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1455
      Left            =   480
      TabIndex        =   14
      Top             =   3360
      Width           =   3135
      Begin VB.CheckBox chk_cardenz 
         Caption         =   "Cardiac Enzyme"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   0
         TabIndex        =   32
         Top             =   120
         Width           =   2055
      End
      Begin VB.CheckBox chk_CPKMB 
         Caption         =   "CPK-MB"
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
         Height          =   375
         Left            =   1920
         TabIndex        =   18
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox chk_CPK 
         Caption         =   "CPK"
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
         Height          =   375
         Left            =   720
         TabIndex        =   17
         Top             =   960
         Width           =   735
      End
      Begin VB.CheckBox chk_LDH 
         Caption         =   "LDH"
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
         Height          =   375
         Left            =   1920
         TabIndex        =   16
         Top             =   480
         Width           =   855
      End
      Begin VB.CheckBox chk_SGOT 
         Caption         =   "SGOT"
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
         Height          =   375
         Left            =   720
         TabIndex        =   15
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5280
      TabIndex        =   10
      Top             =   1800
      Width           =   2895
      Begin VB.CheckBox chk_elect 
         Caption         =   "Electrolytes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   1695
      End
      Begin VB.CheckBox chk_chlor 
         Caption         =   "Chlorides"
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
         Height          =   315
         Left            =   360
         TabIndex        =   13
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox chk_pota 
         Caption         =   "Potassium"
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
         Height          =   375
         Left            =   1440
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
      Begin VB.CheckBox chk_sodium 
         Caption         =   "Sodium"
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
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sugar Profile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1455
      Left            =   480
      TabIndex        =   5
      Top             =   1800
      Width           =   4815
      Begin VB.CheckBox chk_gluco 
         Caption         =   "Glucocheck"
         Enabled         =   0   'False
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
         Height          =   255
         Left            =   2760
         TabIndex        =   9
         Top             =   960
         Width           =   1455
      End
      Begin VB.CheckBox chk_ransug 
         Caption         =   "Random Sugar"
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
         Height          =   255
         Left            =   720
         TabIndex        =   8
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox chk_ppsug 
         Caption         =   "PP Sugar"
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
         Height          =   255
         Left            =   2760
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox chk_FastSug 
         Caption         =   "Fasting Sugar"
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
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.TextBox txt_name 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6720
      TabIndex        =   4
      Top             =   960
      Width           =   4455
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame11 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8280
      TabIndex        =   64
      Top             =   1800
      Width           =   1455
      Begin VB.CheckBox chk_osmolality 
         Caption         =   "Osmolality"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   65
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Haematological Investigations and Biochemistry are mutually Exclusive"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   480
      TabIndex        =   63
      Top             =   6000
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.Label lbl_msg 
      Caption         =   " Select  Procedures upto Eleven only...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   27
      Top             =   8040
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "EMERGENCY LABORATORY SERVICES"
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   4200
      TabIndex        =   26
      Top             =   360
      Width           =   5775
   End
   Begin VB.Label Label1 
      Caption         =   "HOMI BHABHA CANCER HOSPITAL AND RESEARCH CENTRE"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2880
      TabIndex        =   25
      Top             =   0
      Width           =   8655
   End
   Begin VB.Label lbl_name 
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
      Height          =   255
      Left            =   6000
      TabIndex        =   3
      Top             =   960
      Width           =   735
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
      Left            =   1920
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lbl_specimen 
      AutoSize        =   -1  'True
      Caption         =   "Nature of Material"
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
      Left            =   840
      TabIndex        =   0
      Top             =   1440
      Width           =   1860
   End
End
Attribute VB_Name = "Em_lab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hisflag As Boolean
Dim cisflag As Boolean
Dim bloodgasflag As Boolean
Dim obj As Object



Private Sub Check1_Click()

End Sub

Private Sub chk_albumin_Click()
If chk_albumin.Value = 0 Then
    'chk_prot.Value = 0
    chk_livfuntest.Value = 0
End If
If chk_albumin.Value = 1 Then
    chk_prot.Value = 1
End If
If chk_prot.Value = 1 And chk_albumin.Value = 1 Then
    'chk_albumin.Value = 1
    chk_glob.Value = 1
End If
If chk_prot.Value = 0 Or chk_albumin.Value = 0 Then
    'chk_albumin.Value = 1
    chk_glob.Value = 0
End If

End Sub

Private Sub chk_alkphos_Click()
If chk_alkphos.Value = 0 Then
    chk_livfuntest.Value = 0
End If
End Sub

Private Sub chk_amylase_Click()
 If chk_amylase.Value = 1 Then
   hisflag = True
 Else
   hisflag = False
 End If
End Sub

Private Sub chk_bgel_Click()

If chk_bgel.Value = 1 Then
   
   'hisflag = True
   bloodgasflag = True
Else
 
   'hisflag = False
   bloodgasflag = False
End If

End Sub

Private Sub chk_bloodgas_Click()
 If chk_bloodgas.Value = 1 Then
   hisflag = True
 Else
   hisflag = False
 End If
End Sub

Private Sub chk_bt_Click()
 If CHK_BT.Value = 1 Then
   cisflag = True
 Else
   cisflag = False
 End If
End Sub

Private Sub chk_ca_Click()
 If chk_ca.Value = 1 Then
   hisflag = True
 Else
   hisflag = False
 End If
End Sub

Private Sub chk_ca1_Click()
 If chk_ca1.Value = 1 Then
   hisflag = True
 Else
   hisflag = False
 End If
End Sub

Private Sub chk_chlor_Click()
 If chk_chlor.Value = 1 Then
   hisflag = True
 Else
   hisflag = False
 End If
End Sub

Private Sub chk_CPK_Click()
 If CHK_cpk.Value = 1 Then
   hisflag = True
   '***** Shraddha 07/10/2011 *****
    Frame1.Visible = False
    Frame2.Visible = False
    Frame11.Visible = False
    Frame5.Visible = False
    Frame4.Visible = False
'    Frame6.Visible = False
    Frame10.Visible = False
'    Frame8.Visible = False
'***********************
 Else
   hisflag = False
   '***** Shraddha 07/10/2011 *****
    Frame1.Visible = True
    Frame2.Visible = True
    Frame11.Visible = True
    Frame5.Visible = True
    Frame4.Visible = True
'    Frame6.Visible = True
    Frame10.Visible = True
'    Frame8.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
        End If
    Next
'***********************
 End If
End Sub

Private Sub chk_CPKMB_Click()
 If chk_cpkmb.Value = 1 Then
   hisflag = True
   '***** Shraddha 07/10/2011 *****
    Frame1.Visible = False
    Frame2.Visible = False
    Frame11.Visible = False
    Frame5.Visible = False
    Frame4.Visible = False
'    Frame6.Visible = False
    Frame10.Visible = False
'    Frame8.Visible = False
'***********************
 Else
   hisflag = False
   '***** Shraddha 07/10/2011 *****
    Frame1.Visible = True
    Frame2.Visible = True
    Frame11.Visible = True
    Frame5.Visible = True
    Frame4.Visible = True
'    Frame6.Visible = True
    Frame10.Visible = True
'    Frame8.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
        End If
    Next
'***********************
 End If
End Sub

Private Sub chk_creat_Click()
 If chk_creat.Value = 1 Then
   hisflag = True
 Else
   hisflag = False
 End If
End Sub

Private Sub CHK_CT_Click()
 If CHK_CT.Value = 1 Then
   cisflag = True
 Else
   cisflag = False
 End If
End Sub

Private Sub chk_dirbil_Click()

If chk_dirbil.Value = 1 Then
    hisflag = True
Else
    hisflag = False
End If

End Sub

Private Sub chk_FastSug_Click()
 If chk_FastSug.Value = 1 Then
   hisflag = True
 Else
   hisflag = False
 End If
End Sub

Private Sub chk_glob_Click()
If chk_glob.Value = 0 Then
    chk_livfuntest.Value = 0
End If
If chk_prot.Value = 1 And chk_albumin.Value = 1 Then
    'chk_albumin.Value = 1
    chk_glob.Value = 1
End If
If chk_prot.Value = 0 Or chk_albumin.Value = 0 Then
    'chk_albumin.Value = 1
    chk_glob.Value = 0
End If
End Sub

Private Sub chk_gluco_Click()
 If chk_gluco.Value = 1 Then
   hisflag = True
 Else
   hisflag = False
 End If
End Sub

Private Sub chk_hb_Click()
 If CHK_HB.Value = 1 Then
   cisflag = True
 Else
   cisflag = False
 End If
End Sub

Private Sub chk_indirbil_Click()

If chk_indirbil.Value = 1 Then
    hisflag = True
Else
    hisflag = False
End If

End Sub

Private Sub chk_lactate_Click()
 If chk_lactate.Value = 1 Then
   hisflag = True
 Else
   hisflag = False
 End If
End Sub

Private Sub chk_LDH_Click()
 If chk_ldh.Value = 1 Then
   hisflag = True
   '***** Shraddha 07/10/2011 *****
    Frame1.Visible = False
    Frame2.Visible = False
    Frame11.Visible = False
    Frame5.Visible = False
    Frame4.Visible = False
'    Frame6.Visible = False
    Frame10.Visible = False
'    Frame8.Visible = False
'***********************
 Else
   hisflag = False
   '***** Shraddha 07/10/2011 *****
    Frame1.Visible = True
    Frame2.Visible = True
    Frame11.Visible = True
    Frame5.Visible = True
    Frame4.Visible = True
'    Frame6.Visible = True
    Frame10.Visible = True
'    Frame8.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
        End If
    Next
'***********************
 End If
End Sub

Private Sub chk_livfuntest_Click()
If chk_livfuntest.Value = 1 Then
    hisflag = True
    chk_prot.Value = 1
    chk_albumin.Value = 1
    chk_glob.Value = 1
    chk_alkphos.Value = 1
    chk_totbil.Value = 1
    Chk_SGOT1.Value = 1
    chk_sgpt.Value = 1
Else
    hisflag = False
    chk_prot.Value = 0
    chk_albumin.Value = 0
    chk_glob.Value = 0
    chk_alkphos.Value = 0
    chk_totbil.Value = 0
    Chk_SGOT1.Value = 0
    chk_sgpt.Value = 0
End If
End Sub

Private Sub chk_magn_Click()
 If chk_magn.Value = 1 Then
   hisflag = True
 Else
   hisflag = False
 End If
End Sub

Private Sub chk_osmolality_Click()
    
    If chk_osmolality.Value = 1 Then
        hisflag = True
        '***** Shraddha 07/10/2011 *****
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
    Frame5.Visible = False
    Frame4.Visible = False
'    Frame6.Visible = False
    Frame10.Visible = False
'    Frame8.Visible = False
'***********************
    Else
        hisflag = False
'***** Shraddha 07/10/2011 *****
    Frame1.Visible = True
    Frame2.Visible = True
    Frame3.Visible = True
    Frame5.Visible = True
    Frame4.Visible = True
'    Frame6.Visible = True
    Frame10.Visible = True
'    Frame8.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
        End If
    Next
'***********************
    End If
    
End Sub

Private Sub CHK_PC_Click()
 If CHK_PC.Value = 1 Then
   cisflag = True
 Else
   cisflag = False
 End If
End Sub

Private Sub chk_pcv_Click()
  If chk_pcv.Value = 1 Then
   cisflag = True
 Else
   cisflag = False
 End If
End Sub

Private Sub chk_pota_Click()
 If chk_pota.Value = 1 Then
   hisflag = True
 Else
   hisflag = False
End If
End Sub

Private Sub chk_ppsug_Click()
 If chk_ppsug.Value = 1 Then
   hisflag = True
 Else
   hisflag = False
End If
End Sub

Private Sub chk_prot_Click()
If chk_prot.Value = 0 Then
    chk_albumin.Value = 0
    chk_livfuntest.Value = 0
End If
If chk_prot.Value = 1 And chk_albumin.Value = 1 Then
    'chk_albumin.Value = 1
    chk_glob.Value = 1
End If
If chk_prot.Value = 0 Or chk_albumin.Value = 0 Then
    'chk_albumin.Value = 1
    chk_glob.Value = 0
End If

End Sub

Private Sub CHK_PT_Click()
 If chk_pt.Value = 1 Then
   cisflag = True
 Else
   cisflag = False
 End If
End Sub

Private Sub CHK_PTT_Click()
 If CHK_PTT.Value = 1 Then
   cisflag = True
 Else
   cisflag = False
 End If
End Sub

Private Sub chk_ransug_Click()
 If chk_ransug.Value = 1 Then
   hisflag = True
 Else
   hisflag = False
End If
End Sub

Private Sub chk_SGOT_Click()
 If chk_sgot.Value = 1 Then
   hisflag = True
   '***** Shraddha 07/10/2011 *****
    Frame1.Visible = False
    Frame2.Visible = False
    Frame11.Visible = False
    Frame5.Visible = False
    Frame4.Visible = False
'    Frame6.Visible = False
    Frame10.Visible = False
'    Frame8.Visible = False
'***********************
 Else
   hisflag = False
   '***** Shraddha 07/10/2011 *****
    Frame1.Visible = True
    Frame2.Visible = True
    Frame11.Visible = True
    Frame5.Visible = True
    Frame4.Visible = True
'    Frame6.Visible = True
    Frame10.Visible = True
'    Frame8.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
        End If
    Next
'***********************
 End If
End Sub

Private Sub Chk_SGOT1_Click()
If Chk_SGOT1.Value = 0 Then
    chk_livfuntest.Value = 0
End If
End Sub

Private Sub chk_sgpt_Click()
If chk_sgpt.Value = 0 Then
    chk_livfuntest.Value = 0
End If
End Sub

Private Sub chk_am_Click()

If chk_am.Value = 1 Then
   'hisflag = True
   
   bloodgasflag = True
   
 Else
   'hisflag = False
   bloodgasflag = False
End If

'chk_cap.Enabled = False
'chk_venous.Enabled = False

'''' WRITTEN CODE BY MONOMITA DAS
''THEN ADDED BY MILIND
If (chk_cap.Value = 1 And chk_am.Value = 1) Or (chk_venous.Value = 1 And chk_am.Value = 1) Then
        MsgBox "YOU CAN SELECT ONLY ONE !!"
        GoTo MM
ElseIf (chk_am.Value = 0 And chk_cap.Value = 1) Or (chk_am.Value = 0 And chk_venous.Value = 1) Then
        chk_am.Value = 0
        Exit Sub
ElseIf chk_am.Value = 1 And (chk_cap.Value = 0 Or chk_venous.Value = 0) Then
        chk_am.Value = 1
        Exit Sub
End If

MM:
chk_am.Value = 0
chk_venous.Value = 0
chk_cap.Value = 0
End Sub


Private Sub chk_btct_Click()
If chk_btct.Value = 1 Then
    cisflag = True
    CHK_BT.Value = 1
    CHK_CT.Value = 1
Else
    cisflag = False
    CHK_BT.Value = 0
    CHK_CT.Value = 0
End If
End Sub

Private Sub chk_cap_Click()

If chk_cap.Value = 1 Then
   'hisflag = True
   bloodgasflag = True
 Else
   'hisflag = False
   bloodgasflag = False
End If

'chk_am.Enabled = False
'chk_venous.Enabled = False

''''CODE BY MONOMITA DAS
''THEN ADDED BY MILIND

If chk_cap.Value = 1 Then
    chk_bgel.Value = 0
    chk_bgel.Enabled = False

ElseIf chk_cap.Value = 0 Then
    
    chk_bgel.Enabled = True

End If

If (chk_cap.Value = 1 And chk_am.Value = 1) Or (chk_cap.Value = 1 And chk_venous.Value = 1) Then
        MsgBox "YOU CAN SELECT ONLY ONE !!"
        GoTo MM
ElseIf (chk_am.Value = 0 And chk_cap.Value = 1) Or (chk_venous.Value = 0 And chk_cap.Value = 1) Then
        chk_cap.Value = 1
        Exit Sub
ElseIf (chk_am.Value = 1 Or chk_venous.Value = 1) And chk_cap.Value = 0 Then
        chk_cap.Value = 0
        Exit Sub
End If
Exit Sub
MM:

chk_am.Value = 0
chk_venous.Value = 0
chk_cap.Value = 0

End Sub

Private Sub chk_cardenz_Click()

  

If chk_cardenz.Value = 1 Then
'***** Shraddha 07/10/2011 *****
    Frame1.Visible = False
    Frame2.Visible = False
    Frame11.Visible = False
    Frame5.Visible = False
    Frame4.Visible = False
'    Frame6.Visible = False
    Frame10.Visible = False
'    Frame8.Visible = False
'***********************
    hisflag = True
    chk_sgot.Value = 1
    chk_ldh.Value = 1
    'chk_ldh1.Value = 1
    chk_cpkmb.Value = 1
    CHK_cpk.Value = 1
Else

'***** Shraddha 07/10/2011 *****
    Frame1.Visible = True
    Frame2.Visible = True
    Frame11.Visible = True
    Frame5.Visible = True
    Frame4.Visible = True
'    Frame6.Visible = True
    Frame10.Visible = True
'    Frame8.Visible = True
    For Each c In Controls
        If TypeOf c Is CheckBox Then
            c.Value = 0
        End If
    Next
'***********************

    hisflag = False
    chk_sgot.Value = 0
    chk_ldh.Value = 0
    'chk_ldh1.Value = 0
    chk_cpkmb.Value = 0
    CHK_cpk.Value = 0
End If
End Sub

Private Sub chk_elect_Click()
If chk_elect.Value = 1 Then
    hisflag = True
    chk_sodium.Value = 1
    chk_pota.Value = 1
    chk_chlor.Value = 1
Else
    hisflag = False
    chk_sodium.Value = 0
    chk_pota.Value = 0
    chk_chlor.Value = 0
End If
End Sub



Private Sub CHK_hbpcv_Click()
If CHK_hbpcv.Value = 1 Then
    cisflag = True
    CHK_HB.Value = 1
    Chk_wbc.Value = 1
    CHK_PC.Value = 1
Else
    cisflag = False
    CHK_HB.Value = 0
    Chk_wbc.Value = 0
    CHK_PC.Value = 0
End If
End Sub

Private Sub chk_sodium_Click()
 If chk_sodium.Value = 1 Then
   hisflag = True
 Else
   hisflag = False
 End If
End Sub

Private Sub chk_totbil_Click()
If chk_totbil.Value = 0 Then
    chk_livfuntest.Value = 0
End If
End Sub

Private Sub CHK_tt_Click()
 If CHK_TT.Value = 1 Then
   cisflag = True
 Else
   cisflag = False
 End If
End Sub

Private Sub chk_ur24hr_Click()
    If chk_ur24hr.Value = 1 Then
        hisflag = True
    Else
        hisflag = False
    End If
End Sub

Private Sub chk_urea_Click()
 If chk_urea.Value = 1 Then
   hisflag = True
   'bloodgasflag = True
 Else
   hisflag = False
   'bloodgasflag = False
   
 End If
End Sub

Private Sub chk_uricacid_Click()
    If chk_uricacid.Value = 1 Then
        hisflag = True
    Else
        hisflag = False
    End If
End Sub

Private Sub chk_urrandom_Click()
If chk_urrandom.Value = 1 Then
    hisflag = True
Else
    hisflag = False
End If
End Sub

Private Sub chk_venous_Click()

If chk_venous.Value = 1 Then
        
        'hisflag = True
        bloodgasflag = True
    Else
        
        'hisflag = False
        bloodgasflag = False

End If

'chk_am.Enabled = False
'chk_cap.Enabled = False

If (chk_venous.Value = 1 And chk_am.Value = 1) Or (chk_cap.Value = 1 And chk_cap.Value = 1) Then
        MsgBox "YOU CAN SELECT ONLY ONE !!"
        GoTo MM
ElseIf (chk_am.Value = 0 And chk_venous.Value = 1) Or (chk_venous.Value = 1 And chk_cap.Value = 1) Then
        chk_venous.Value = 1
        Exit Sub
ElseIf (chk_am.Value = 1 Or chk_cap.Value = 1) And chk_venous.Value = 0 Then
        chk_venous.Value = 0
        Exit Sub
End If
Exit Sub
MM:
chk_am.Value = 0
chk_venous.Value = 0
chk_cap.Value = 0


End Sub

Private Sub Chk_wbc_Click()
 
 If Chk_wbc.Value = 1 Then
   
   cisflag = True
 
 Else
   
   cisflag = False
 
 End If
 
End Sub

Private Sub CMB_NMETERIAL_Click()

Dim ctrl As Object
For Each ctrl In Controls
    If TypeOf ctrl Is CheckBox Then
        ctrl.Value = 0
    End If
Next

    If Trim(CMB_NMETERIAL.Text) <> "Blood" Then
        
        If Trim(CMB_NMETERIAL.Text) = "Urine ( 24 Hours )" Then
            Frame2.Visible = False
            Frame10.Visible = False
            chk_ur24hr.Visible = True
            chk_urrandom.Visible = False
            Frame5.Visible = True
            Frame5.Enabled = True
            Frame5.Top = 3000
            Frame5.Left = 4500
            Frame5.Width = 4500
            
        ElseIf Trim(CMB_NMETERIAL.Text) = "Urine ( Random )" Then
            Frame2.Visible = False
            Frame10.Visible = False
            chk_ur24hr.Visible = False
            chk_urrandom.Visible = True
            '***** Added by Shraddha on 20/07/10 *****
            chk_urea.Visible = False
            chk_creat.Visible = False
            '*****************************************
            Frame5.Visible = True
            Frame5.Enabled = True
            Frame5.Top = 3000
            Frame5.Left = 4500
            Frame5.Width = 4500

        ElseIf Trim(CMB_NMETERIAL.Text) = "Fluid" Then
            Frame2.Visible = True
            Frame10.Visible = False
            chk_ur24hr.Visible = False
            chk_urrandom.Visible = False
            Frame5.Top = 1800
            Frame5.Left = 8840
            Frame5.Width = 2000
        End If
        
        Frame1.Visible = False
        Frame3.Visible = False
'        Frame6.Visible = False
        Frame11.Visible = False
        Frame4.Visible = False
        Frame7.Visible = False
        Label3.Visible = False
        chk_ca1.Visible = False
        chk_magn.Visible = False
        chk_ca.Visible = False
        chk_amylase.Visible = False
        chk_lactate.Visible = False
    
    Else
    
        Frame2.Visible = True
        '***** Shraddha 26/07/10 *****
        'If 1st floor recovery or ICU then visible frame LFT
        'If ipadd = "150.150.200.62" Or ipadd = "150.150.15.51" Then
            'Frame10.Visible = True
        'End If
        'If ipadd = "10.3.1.56" Then
            'Frame10.Visible = True
        'End If
        'Again commented by Shraddha on 07/10/2011
        'Made it open to all
        'Shraddha 10/09/10
'        Dim icu_QR As rdoQuery
'        Dim icu_RS As rdoResultset
'        Dim sql As String
'
'        sql = "SELECT * FROM DISLIB.disicu"
'        Set icu_QR = rdb.CreateQuery("", sql)
'        Set icu_RS = icu_QR.OpenResultset(1, 3)
'
'        While Not icu_RS.EOF
'            If Trim(icu_RS("IPADD")) = ipadd Then
'                Frame10.Visible = True
'            End If
'            icu_RS.MoveNext
'        Wend
'        '*****************************
        Frame10.Visible = True
        Frame1.Visible = True
        Frame3.Visible = True
        'Frame6.Visible = True 'Shraddha 24/10/08
       ' Frame11.Visible = True 'Shraddha 24/10/08
        Frame4.Visible = True
        'Frame7.Visible = True 'Shraddha 02/07/10
        'Label3.Visible = True 'Shraddha 02/07/10
        chk_ca1.Visible = True
        chk_magn.Visible = True
        'chk_ca.Visible = True
        chk_amylase.Visible = True
        chk_lactate.Visible = True
        chk_ur24hr.Visible = False
        chk_urrandom.Visible = False
        Frame5.Width = 2000
        Frame5.Top = 1560
        Frame5.Left = 9840
          
        
            Dim CTRL1 As Object

        For Each CTRL1 In Controls

            If TypeOf CTRL1 Is Frame Then
            CTRL1.Enabled = True
        End If
        Next

    End If
        
End Sub

Private Sub cmb_nmeterial_GotFocus()
CMB_NMETERIAL.clear
CMB_NMETERIAL.AddItem "Blood"
'CMB_NMETERIAL.AddItem "Urine ( 24 Hours )" 'Again commented on 14/09/2011 Added on 13/07/11
CMB_NMETERIAL.AddItem "Urine ( Random )"
'CMB_NMETERIAL.AddItem "Fluid" 'Commented by Shraddha on 20/07/10
End Sub

Private Sub cmd_back_Click()

If hisflag And cisflag Then
    MsgBox " Heamatological Invesigations and Biochemistry are mutually exclusive, separate requisitions required ", vbInformation
    Exit Sub
ElseIf bloodgasflag And cisflag Then
    MsgBox " Heamatological Invesigations and Biochemistry are mutually exclusive, separate requisitions required ", vbInformation
    Exit Sub
ElseIf hisflag And bloodgasflag Then
    MsgBox "Blood gas Investigations and Other Biochemistry Investigations are mutually exclusive, separate requisitions required", vbInformation
    Exit Sub
End If

If pathology.CMB_TEXT.ListCount <> 0 Then
 pathology.CMB_TEXT.clear
End If
If ELFLAG = False Then
    ccount = 0
    pathology.CMB_TEXT.clear
    '**** SUGAR PROFILE
    If chk_FastSug.Value = 1 Then
        If ccount >= 11 Then
            MsgBox "You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        
    'pathology.CMB_TEXT.AddItem "F842"
     pathology.CMB_TEXT.AddItem "F979"
     ccount = ccount + 1
    End If
    If chk_ppsug.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
    'pathology.CMB_TEXT.AddItem "F843"
    pathology.CMB_TEXT.AddItem "F980"
    ccount = ccount + 1
    End If
    If chk_ransug.Value = 1 Then
        If ccount >= 11 Then
            MsgBox "You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
    'pathology.CMB_TEXT.AddItem "F841"
    pathology.CMB_TEXT.AddItem "F978"
    ccount = ccount + 1
    End If
    If chk_gluco.Value = 1 Then
        If ccount >= 11 Then
            MsgBox "You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
    pathology.CMB_TEXT.AddItem "F848"
    ccount = ccount + 1
    End If
    '**** SUGAR PROFILE
    '**** eLECTROLYTES
    If chk_elect.Value = 1 Then
        If ccount >= 11 Then
            MsgBox "You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        'pathology.CMB_TEXT.AddItem "F860"
        pathology.CMB_TEXT.AddItem "F981"
        ccount = ccount + 1
    Else
        If chk_sodium.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            'pathology.CMB_TEXT.AddItem "F861"
             pathology.CMB_TEXT.AddItem "F982"
            ccount = ccount + 1
        End If
        If chk_pota.Value = 1 Then
            If ccount >= 11 Then
                MsgBox " You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
        'pathology.CMB_TEXT.AddItem "F862"
        pathology.CMB_TEXT.AddItem "F983"
        ccount = ccount + 1
        End If
        If chk_chlor.Value = 1 Then
            If ccount >= 11 Then
                MsgBox " You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
        'pathology.CMB_TEXT.AddItem "F863"
        pathology.CMB_TEXT.AddItem "F984"
        ccount = ccount + 1
        End If
    End If
    '**** Electrolytes
    If chk_urea.Value = 1 Then
        If ccount >= 11 Then
            MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        'pathology.CMB_TEXT.AddItem "F855"
        pathology.CMB_TEXT.AddItem "F985"
        ccount = ccount + 1
     End If
     If chk_creat.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        'pathology.CMB_TEXT.AddItem "F857"
        pathology.CMB_TEXT.AddItem "F987"
        ccount = ccount + 1
    End If
    
    '''''''ADDED BY KANCHAN
    
    If chk_urrandom.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        'pathology.CMB_TEXT.AddItem "F857"
        pathology.CMB_TEXT.AddItem "F031"
        ccount = ccount + 1
    End If
    
    If chk_ur24hr.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        'pathology.CMB_TEXT.AddItem "F857"
        pathology.CMB_TEXT.AddItem "F032"
        ccount = ccount + 1
    End If
    '''''''''''KANCHAN
    '**** Cardiac enzymes
    If chk_cardenz.Value = 1 Then
        If ccount >= 11 Then
            MsgBox "You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        pathology.CMB_TEXT.AddItem "F875"
        ccount = ccount + 1
    Else
        If chk_sgot.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F877"
            ccount = ccount + 1
        End If
    '    If chk_ldh1.Value = 1 Then
    '        If ccount >= 11 Then
    '            MsgBox "You Can Select Max. of 11 Procedures only."
    '            Unload Me
    '            pathology.Show
    '            Exit Sub
    '        End If
    '    pathology.CMB_TEXT.AddItem "F875"
    '    ccount = ccount + 1
    '    End If
        If chk_ldh.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            ccount = ccount + 1
            pathology.CMB_TEXT.AddItem "F914"
        End If
        If CHK_cpk.Value = 1 Then
            If ccount >= 11 Then
                MsgBox " You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            ccount = ccount + 1
            pathology.CMB_TEXT.AddItem "F878"
        End If
        If chk_cpkmb.Value = 1 Then
            If ccount >= 11 Then
                MsgBox " You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
        pathology.CMB_TEXT.AddItem "F879"
        ccount = ccount + 1
        End If
    End If
    '**** cardiac enzymes
    If chk_magn.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        'pathology.CMB_TEXT.AddItem "F891"
        pathology.CMB_TEXT.AddItem "F993"
        ccount = ccount + 1
    End If
    If chk_lactate.Value = 1 Then
        If ccount >= 11 Then
            MsgBox "You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        'pathology.CMB_TEXT.AddItem "F892"
        pathology.CMB_TEXT.AddItem "F990"
        ccount = ccount + 1
    End If
    If chk_amylase.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        'pathology.CMB_TEXT.AddItem "F881"
        pathology.CMB_TEXT.AddItem "F988"
    End If
    If chk_uricacid.Value = 1 Then
        If ccount >= 11 Then
            MsgBox " You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        'pathology.CMB_TEXT.AddItem "F881"
        pathology.CMB_TEXT.AddItem "F986"
    End If
    
    'shweta for 2 new test 'lipase phosphorus
     If chk_phosphorus.Value = 1 Then
        If ccount >= 11 Then
            MsgBox "You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
    
    pathology.CMB_TEXT.AddItem "F994"
    ccount = ccount + 1
    End If
    
     If chk_lipase.Value = 1 Then
        If ccount >= 11 Then
            MsgBox "You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
    
    pathology.CMB_TEXT.AddItem "F989"
    ccount = ccount + 1
    End If
    
    
'    If chk_osmolality.Value = 1 Then
'
'        If ccount >= 11 Then
'            MsgBox " You Can Select Max. of 11 Procedures only."
'            Unload Me
'            pathology.show
'            Exit Sub
'        End If
'
'        If count > 1 Then
'
'        Else
'            ccount = ccount + 1
'            'pathology.CMB_TEXT.AddItem "F881"
'            pathology.CMB_TEXT.AddItem "F998"
'        End If
'
'    End If
    
    If chk_ca1.Value = 1 Then
        If ccount >= 11 Then
            MsgBox "You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        'pathology.CMB_TEXT.AddItem "F888"
        pathology.CMB_TEXT.AddItem "F991"
    End If
    If chk_ca.Value = 1 Then
        If ccount >= 11 Then
            MsgBox "You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        
        'pathology.CMB_TEXT.AddItem "F889"
        pathology.CMB_TEXT.AddItem "F992"
    End If
    
    '''****** ADD BY MONO ON 09/04/2003
    '**** Blood Gas
    '***** ARTERIAL
    
    If chk_am.Value = 1 Then
        If ccount >= 11 Then
            MsgBox "You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        pathology.CMB_TEXT.AddItem "F901"
    End If
    
    '**** VENOUS
    If chk_venous.Value = 1 Then
        If ccount >= 11 Then
            MsgBox "You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        pathology.CMB_TEXT.AddItem "F995"
    End If
    '****
    
    '***** CAPILLARY
    If chk_cap.Value = 1 Then
        If ccount >= 11 Then
            MsgBox "You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        pathology.CMB_TEXT.AddItem "F902"
    End If
    
    If chk_bgel.Value = 1 Then
        If ccount >= 11 Then
            MsgBox "You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        pathology.CMB_TEXT.AddItem "F022"
    End If
    '**** Blood gas
    '**** Emergency Lab for Heamatology
    '**** hp and pcv
    If CHK_hbpcv.Value = 1 Then
        If ccount >= 11 Then
            MsgBox "You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        pathology.CMB_TEXT.AddItem "F903"
    Else
        If CHK_HB.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            ccount = ccount + 1
            pathology.CMB_TEXT.AddItem "F904"
        End If
        
        If Chk_wbc.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            ccount = ccount + 1
            pathology.CMB_TEXT.AddItem "F905"
        End If
        
        If CHK_PC.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            ccount = ccount + 1
            pathology.CMB_TEXT.AddItem "F906"
        End If
    End If
    
    '**** hb and pcv
    '**** bt and ct
    If chk_btct.Value = 1 Then
        If ccount >= 11 Then
            MsgBox "You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        pathology.CMB_TEXT.AddItem "F908"
    Else
        If CHK_BT.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            ccount = ccount + 1
            pathology.CMB_TEXT.AddItem "F909"
        End If
        If CHK_CT.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            ccount = ccount + 1
            pathology.CMB_TEXT.AddItem "F910"
        End If
    End If
    '**** bt and ct
    '**** pcv
    If chk_pcv.Value = 1 Then
        If ccount >= 11 Then
            MsgBox "You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        pathology.CMB_TEXT.AddItem "F907"
    End If
    '**** pcv
    If chk_pt.Value = 1 Then
        If ccount >= 11 Then
            MsgBox "You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        pathology.CMB_TEXT.AddItem "F911"
    End If
    If CHK_PTT.Value = 1 Then
        If ccount >= 11 Then
            MsgBox "You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        pathology.CMB_TEXT.AddItem "F912"
    End If
    If CHK_TT.Value = 1 Then
        If ccount >= 11 Then
            MsgBox "You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        ccount = ccount + 1
        pathology.CMB_TEXT.AddItem "F913"
    End If
    '**** Emergency lab for Heamatology
    '**** Liver function Tests
    If chk_livfuntest.Value = 1 Then
        If ccount > 11 Then
            MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
            Unload Me
            pathology.show
            Exit Sub
        End If
        pathology.CMB_TEXT.AddItem "F949"
        ccount = ccount + 1
        If chk_dirbil.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F959"
            ccount = ccount + 1
        End If
        If chk_indirbil.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F960"
            ccount = ccount + 1
        End If
    Else
        If chk_prot.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F950"
            ccount = ccount + 1
        End If
        If chk_albumin.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F955"
            ccount = ccount + 1
        End If
        If chk_glob.Value = 1 Then
            If ccount > 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F956"
            ccount = ccount + 1
        End If
        If chk_alkphos.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F957"
            ccount = ccount + 1
        End If
        If chk_totbil.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F958"
            ccount = ccount + 1
        End If
        If chk_dirbil.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F959"
            ccount = ccount + 1
        End If
        If chk_indirbil.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F960"
            ccount = ccount + 1
        End If
        
        If Chk_SGOT1.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F877"
            ccount = ccount + 1
        End If
        If chk_sgpt.Value = 1 Then
            If ccount >= 11 Then
                MsgBox "Deselect the Procedures, You Can Select Max. of 11 Procedures only."
                Unload Me
                pathology.show
                Exit Sub
            End If
            pathology.CMB_TEXT.AddItem "F961"
            ccount = ccount + 1
        End If
    End If
    '**** Liver function Tests
    
    '***** Shraddha 19/07/2011 *****
    If chk_osmolality.Value = 1 Then
        
        If ccount >= 1 Then
            MsgBox "Please make separate requisition for Osmolality"
            Exit Sub
        Else
            ccount = ccount + 1
            pathology.CMB_TEXT.AddItem "F998"
        End If
            
    End If
    '*****
End If

meterial = Trim(CMB_NMETERIAL.Text)

Unload Me
pathology.show
End Sub

Private Sub txt_nmeterial_LostFocus()
meterial = Trim(CMB_NMETERIAL.Text)
End Sub

Private Sub Form_Load()
pathology.CMB_TEXT.clear
cisflag = False  ' on 03/02
hisflag = False
bloodgasflag = False

Dim ctrl As Object

For Each ctrl In Controls

    If TypeOf ctrl Is Frame Then
        ctrl.Enabled = False
    End If
Next
End Sub

