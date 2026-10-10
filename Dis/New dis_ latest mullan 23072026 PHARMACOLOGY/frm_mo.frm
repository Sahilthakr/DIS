VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_mo 
   Caption         =   "Form3"
   ClientHeight    =   9705
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13170
   LinkTopic       =   "Form3"
   ScaleHeight     =   9705
   ScaleWidth      =   13170
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Winsock1 
      Height          =   480
      Left            =   12600
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   126
      Top             =   8040
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "Back to BatchCommit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   440
      Left            =   9000
      MouseIcon       =   "frm_mo.frx":0000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   125
      Top             =   9120
      Width           =   2535
   End
   Begin VB.Frame common 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   480
      TabIndex        =   13
      Top             =   840
      Width           =   12255
      Begin VB.TextBox txt_DMG 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   6960
         TabIndex        =   25
         Top             =   1155
         Width           =   4095
      End
      Begin VB.TextBox txt_pathno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   10320
         TabIndex        =   24
         Top             =   315
         Width           =   1695
      End
      Begin VB.ComboBox cmb_reqno 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   6600
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txt_name 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2400
         TabIndex        =   22
         Top             =   720
         Width           =   5775
      End
      Begin VB.TextBox txt_sex 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2400
         TabIndex        =   21
         Top             =   1155
         Width           =   615
      End
      Begin VB.TextBox txt_age 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3600
         TabIndex        =   20
         Top             =   1155
         Width           =   735
      End
      Begin VB.TextBox txt_category 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   5400
         TabIndex        =   19
         Top             =   1155
         Width           =   975
      End
      Begin VB.ComboBox cmb_reportdoc 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   0
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         Height          =   975
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1335
         Begin VB.OptionButton Opt_patient 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Patient"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton opt_staff 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Staff"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.TextBox txt_nmaterial 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   9360
         TabIndex        =   14
         Top             =   720
         Width           =   2775
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   2400
         TabIndex        =   26
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         ClipMode        =   1
         ForeColor       =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label lbl_dmg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DMG"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6480
         TabIndex        =   35
         Top             =   1200
         Width           =   435
      End
      Begin VB.Label lbl_pathno 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PathNo."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   9480
         TabIndex        =   34
         Top             =   360
         Width           =   780
      End
      Begin VB.Label lbl_caseno 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Case No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1560
         TabIndex        =   33
         Top             =   360
         Width           =   885
      End
      Begin VB.Label lbl_reqno 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Requisition No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5040
         TabIndex        =   32
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lbl_name 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1800
         TabIndex        =   31
         Top             =   720
         Width           =   540
      End
      Begin VB.Label lbl_sex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1560
         TabIndex        =   30
         Top             =   1200
         Width           =   690
      End
      Begin VB.Label lbl_age 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3120
         TabIndex        =   29
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lbl_category 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4440
         TabIndex        =   28
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Specimen"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   8400
         TabIndex        =   27
         Top             =   720
         Width           =   945
      End
   End
   Begin VB.Frame COMMAND 
      BackColor       =   &H00FFC0C0&
      Height          =   1095
      Left            =   480
      TabIndex        =   3
      Top             =   8520
      Width           =   12135
      Begin VB.CommandButton CMD_NEW 
         BackColor       =   &H00FF8080&
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   440
         Left            =   120
         MouseIcon       =   "frm_mo.frx":030A
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton cmd_view 
         BackColor       =   &H00FF8080&
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   440
         Left            =   1320
         MouseIcon       =   "frm_mo.frx":0614
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton cmd_modify 
         BackColor       =   &H00FF8080&
         Caption         =   "Modify"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   440
         Left            =   2520
         MouseIcon       =   "frm_mo.frx":091E
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton cmd_save 
         BackColor       =   &H00FF8080&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   440
         Left            =   3720
         MouseIcon       =   "frm_mo.frx":0C28
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton cmd_cancel 
         BackColor       =   &H00FF8080&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   440
         Left            =   4920
         MouseIcon       =   "frm_mo.frx":0F32
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00FF8080&
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   440
         Left            =   7440
         MouseIcon       =   "frm_mo.frx":123C
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton cmd_back 
         BackColor       =   &H00FF8080&
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8640
         MouseIcon       =   "frm_mo.frx":1546
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmd_commit 
         BackColor       =   &H00FF8080&
         Caption         =   "Commit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   440
         Left            =   6120
         MouseIcon       =   "frm_mo.frx":1850
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmd_exit 
         BackColor       =   &H00FF8080&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   440
         Left            =   9600
         MouseIcon       =   "frm_mo.frx":1B5A
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.TextBox txt_MOremarks 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1560
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   7440
      Width           =   11295
   End
   Begin VB.ComboBox cmb_finalby 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6960
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   8160
      Width           =   3615
   End
   Begin VB.ComboBox cmb_enteryby 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   8160
      Width           =   3735
   End
   Begin VB.PictureBox WIN_CLIENT 
      Height          =   480
      Left            =   360
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   127
      Top             =   120
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab_MO 
      Height          =   4695
      Left            =   480
      TabIndex        =   36
      Top             =   2640
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   8281
      _Version        =   393216
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "G408: Respiratory"
      TabPicture(0)   =   "frm_mo.frx":1E64
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label15"
      Tab(0).Control(1)=   "Label14"
      Tab(0).Control(2)=   "cmb_408BactDet"
      Tab(0).Control(3)=   "cmb_408virDet"
      Tab(0).Control(4)=   "Txt_BacteriaG408"
      Tab(0).Control(5)=   "cmd_addB408"
      Tab(0).Control(6)=   "Txt_VirusG408"
      Tab(0).Control(7)=   "cmd_AddV408"
      Tab(0).Control(8)=   "ListBacteriaG408"
      Tab(0).Control(9)=   "ListVirusG408"
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "G409: Meningitis Panel"
      TabPicture(1)   =   "frm_mo.frx":1E80
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Txt_YeastG409"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmd_addY409"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "ListYeastG409"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmb_409YeastDet"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "ListVirusG409"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "ListBacteriaG409"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmd_AddV409"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Txt_VirusG409"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "cmd_addB409"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Txt_BacteriaG409"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "cmb_409virDet"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "cmb_409BactDet"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "G406: Gastro-intestinal"
      TabPicture(2)   =   "frm_mo.frx":1E9C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmb_406BactDet"
      Tab(2).Control(1)=   "cmb_406virDet"
      Tab(2).Control(2)=   "Txt_BacteriaG406"
      Tab(2).Control(3)=   "cmd_addB406"
      Tab(2).Control(4)=   "Txt_VirusG406"
      Tab(2).Control(5)=   "cmd_AddV406"
      Tab(2).Control(6)=   "ListBacteriaG406"
      Tab(2).Control(7)=   "ListVirusG406"
      Tab(2).Control(8)=   "cmb_406ParaDet"
      Tab(2).Control(9)=   "ListParaG406"
      Tab(2).Control(10)=   "cmd_addP406"
      Tab(2).Control(11)=   "Txt_ParaG406"
      Tab(2).Control(12)=   "Label8"
      Tab(2).Control(13)=   "Label6"
      Tab(2).Control(14)=   "Label5"
      Tab(2).ControlCount=   15
      TabCaption(3)   =   "G410: Pneumonia"
      TabPicture(3)   =   "frm_mo.frx":1EB8
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label12"
      Tab(3).Control(1)=   "Label11"
      Tab(3).Control(2)=   "Label10"
      Tab(3).Control(3)=   "Label9"
      Tab(3).Control(4)=   "Txt_BacteriaG410"
      Tab(3).Control(5)=   "Txt_AtypicalG410"
      Tab(3).Control(6)=   "cmd_AddAtypical410"
      Tab(3).Control(7)=   "ListAtypicalG410"
      Tab(3).Control(8)=   "cmb_410AtypicalDet"
      Tab(3).Control(9)=   "Txt_AntiG410"
      Tab(3).Control(10)=   "cmd_AddAnti410"
      Tab(3).Control(11)=   "ListAntiG410"
      Tab(3).Control(12)=   "cmb_410AntiDet"
      Tab(3).Control(13)=   "ListVirusG410"
      Tab(3).Control(14)=   "ListBacteriaG410"
      Tab(3).Control(15)=   "cmd_AddV410"
      Tab(3).Control(16)=   "Txt_VirusG410"
      Tab(3).Control(17)=   "cmd_addB410"
      Tab(3).Control(18)=   "cmb_410virDet"
      Tab(3).Control(19)=   "cmb_410BactDet"
      Tab(3).ControlCount=   20
      TabCaption(4)   =   "G407: Blood Culture"
      TabPicture(4)   =   "frm_mo.frx":1ED4
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmd_Go"
      Tab(4).Control(1)=   "cmb_ACTIVEYN"
      Tab(4).Control(2)=   "Txt_407Gneg"
      Tab(4).Control(3)=   "Txt_407Gpos"
      Tab(4).Control(4)=   "cmb_407GramNeg"
      Tab(4).Control(5)=   "cmb_407GramPos"
      Tab(4).Control(6)=   "cmd_Add407Gneg"
      Tab(4).Control(7)=   "cmd_Add407Gpos"
      Tab(4).Control(8)=   "ListG407Gneg"
      Tab(4).Control(9)=   "ListG407Gpos"
      Tab(4).Control(10)=   "cmb_407Anti"
      Tab(4).Control(11)=   "ListG407Anti"
      Tab(4).Control(12)=   "cmd_Add407Ganti"
      Tab(4).Control(13)=   "Txt_407Anti"
      Tab(4).Control(14)=   "cmb_407Yeast"
      Tab(4).Control(15)=   "ListG407Yeast"
      Tab(4).Control(16)=   "cmd_Add407Gyeast"
      Tab(4).Control(17)=   "Txt_407Yeast"
      Tab(4).Control(18)=   "Label20"
      Tab(4).Control(19)=   "Label19"
      Tab(4).Control(20)=   "Label18"
      Tab(4).Control(21)=   "Label17"
      Tab(4).Control(22)=   "Label13"
      Tab(4).ControlCount=   23
      Begin VB.ListBox ListVirusG408 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Left            =   -74880
         Style           =   1  'Checkbox
         TabIndex        =   102
         Top             =   960
         Width           =   6000
      End
      Begin VB.ListBox ListBacteriaG408 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Left            =   -68520
         Style           =   1  'Checkbox
         TabIndex        =   101
         Top             =   960
         Width           =   5500
      End
      Begin VB.CommandButton cmd_AddV408 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Add Virus"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73320
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   3360
         Width           =   2175
      End
      Begin VB.TextBox Txt_VirusG408 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   99
         Top             =   3840
         Width           =   6000
      End
      Begin VB.CommandButton cmd_addB408 
         BackColor       =   &H0080FF80&
         Caption         =   "Add Bacteria"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -66600
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   3360
         Width           =   1575
      End
      Begin VB.TextBox Txt_BacteriaG408 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -68520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   97
         Top             =   3840
         Width           =   5500
      End
      Begin VB.ComboBox cmb_408virDet 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "frm_mo.frx":1EF0
         Left            =   -74040
         List            =   "frm_mo.frx":1F00
         Style           =   2  'Dropdown List
         TabIndex        =   96
         Top             =   480
         Width           =   3015
      End
      Begin VB.ComboBox cmb_408BactDet 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "frm_mo.frx":1F29
         Left            =   -67440
         List            =   "frm_mo.frx":1F39
         Style           =   2  'Dropdown List
         TabIndex        =   95
         Top             =   480
         Width           =   3015
      End
      Begin VB.ComboBox cmb_409BactDet 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "frm_mo.frx":1F62
         Left            =   4920
         List            =   "frm_mo.frx":1F72
         Style           =   2  'Dropdown List
         TabIndex        =   94
         Top             =   480
         Width           =   3255
      End
      Begin VB.ComboBox cmb_409virDet 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "frm_mo.frx":1F9B
         Left            =   720
         List            =   "frm_mo.frx":1FAB
         Style           =   2  'Dropdown List
         TabIndex        =   93
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox Txt_BacteriaG409 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   92
         Top             =   3840
         Width           =   4000
      End
      Begin VB.CommandButton cmd_addB409 
         BackColor       =   &H0080FF80&
         Caption         =   "Add Bacteria"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox Txt_VirusG409 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   90
         Top             =   3840
         Width           =   3900
      End
      Begin VB.CommandButton cmd_AddV409 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Add Virus"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   3360
         Width           =   1695
      End
      Begin VB.ListBox ListBacteriaG409 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Left            =   4200
         Style           =   1  'Checkbox
         TabIndex        =   88
         Top             =   960
         Width           =   4000
      End
      Begin VB.ListBox ListVirusG409 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   87
         Top             =   960
         Width           =   3900
      End
      Begin VB.ComboBox cmb_409YeastDet 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "frm_mo.frx":1FD4
         Left            =   9000
         List            =   "frm_mo.frx":1FE4
         Style           =   2  'Dropdown List
         TabIndex        =   86
         Top             =   480
         Width           =   3255
      End
      Begin VB.ListBox ListYeastG409 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Left            =   8400
         Style           =   1  'Checkbox
         TabIndex        =   85
         Top             =   960
         Width           =   3770
      End
      Begin VB.CommandButton cmd_addY409 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Add Yeast"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox Txt_YeastG409 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   83
         Top             =   3840
         Width           =   3770
      End
      Begin VB.TextBox Txt_ParaG406 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -66600
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   82
         Top             =   3840
         Width           =   3770
      End
      Begin VB.CommandButton cmd_addP406 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Add Parasites"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -65880
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   3360
         Width           =   1815
      End
      Begin VB.ListBox ListParaG406 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Left            =   -66600
         Style           =   1  'Checkbox
         TabIndex        =   80
         Top             =   960
         Width           =   3770
      End
      Begin VB.ComboBox cmb_406ParaDet 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "frm_mo.frx":200D
         Left            =   -65760
         List            =   "frm_mo.frx":201D
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Top             =   480
         Width           =   2895
      End
      Begin VB.ListBox ListVirusG406 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Left            =   -74880
         Style           =   1  'Checkbox
         TabIndex        =   78
         Top             =   960
         Width           =   3900
      End
      Begin VB.ListBox ListBacteriaG406 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Left            =   -70800
         Style           =   1  'Checkbox
         TabIndex        =   77
         Top             =   960
         Width           =   4000
      End
      Begin VB.CommandButton cmd_AddV406 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Add Virus"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73920
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox Txt_VirusG406 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   75
         Top             =   3840
         Width           =   3900
      End
      Begin VB.CommandButton cmd_addB406 
         BackColor       =   &H0080FF80&
         Caption         =   "Add Bacteria"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69720
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox Txt_BacteriaG406 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -70800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   73
         Top             =   3840
         Width           =   4000
      End
      Begin VB.ComboBox cmb_406virDet 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "frm_mo.frx":2046
         Left            =   -74040
         List            =   "frm_mo.frx":2056
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Top             =   480
         Width           =   3015
      End
      Begin VB.ComboBox cmb_406BactDet 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "frm_mo.frx":207F
         Left            =   -70080
         List            =   "frm_mo.frx":208F
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   480
         Width           =   3255
      End
      Begin VB.ComboBox cmb_410BactDet 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "frm_mo.frx":20B8
         Left            =   -67800
         List            =   "frm_mo.frx":20C8
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   360
         Width           =   3615
      End
      Begin VB.ComboBox cmb_410virDet 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "frm_mo.frx":20F1
         Left            =   -73800
         List            =   "frm_mo.frx":2101
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   360
         Width           =   4815
      End
      Begin VB.CommandButton cmd_addB410 
         BackColor       =   &H0080FF80&
         Caption         =   "Add Bacteria"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -68880
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Txt_VirusG410 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   67
         Top             =   2040
         Width           =   6135
      End
      Begin VB.CommandButton cmd_AddV410 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Add Virus"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   1200
         Width           =   975
      End
      Begin VB.ListBox ListBacteriaG410 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   -67800
         Style           =   1  'Checkbox
         TabIndex        =   65
         Top             =   840
         Width           =   4935
      End
      Begin VB.ListBox ListVirusG410 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   -73800
         Style           =   1  'Checkbox
         TabIndex        =   64
         Top             =   840
         Width           =   4815
      End
      Begin VB.ComboBox cmb_410AntiDet 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "frm_mo.frx":212A
         Left            =   -72120
         List            =   "frm_mo.frx":213A
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   2640
         Width           =   3135
      End
      Begin VB.ListBox ListAntiG410 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   -73800
         Style           =   1  'Checkbox
         TabIndex        =   62
         Top             =   3120
         Width           =   4815
      End
      Begin VB.CommandButton cmd_AddAnti410 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Add Anti microbial genes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox Txt_AntiG410 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   60
         Top             =   4080
         Width           =   6135
      End
      Begin VB.ComboBox cmb_410AtypicalDet 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "frm_mo.frx":2163
         Left            =   -67200
         List            =   "frm_mo.frx":2173
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   2640
         Width           =   3855
      End
      Begin VB.ListBox ListAtypicalG410 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   -67800
         Style           =   1  'Checkbox
         TabIndex        =   58
         Top             =   3120
         Width           =   4935
      End
      Begin VB.CommandButton cmd_AddAtypical410 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Add Atypical Bacteria"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -68880
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox Txt_AtypicalG410 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   -68760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   56
         Top             =   4080
         Width           =   6015
      End
      Begin VB.TextBox Txt_407Yeast 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   -67800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   55
         Top             =   4080
         Width           =   4935
      End
      Begin VB.CommandButton cmd_Add407Gyeast 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Add Yeast"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -68880
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   3240
         Width           =   975
      End
      Begin VB.ListBox ListG407Yeast 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   -67800
         Style           =   1  'Checkbox
         TabIndex        =   53
         Top             =   3120
         Width           =   4935
      End
      Begin VB.ComboBox cmb_407Yeast 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "frm_mo.frx":219C
         Left            =   -67800
         List            =   "frm_mo.frx":21AC
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox Txt_407Anti 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   -73800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   51
         Top             =   4080
         Width           =   4815
      End
      Begin VB.CommandButton cmd_Add407Ganti 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Add Antibiotic genes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   3240
         Width           =   975
      End
      Begin VB.ListBox ListG407Anti 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   -73800
         Style           =   1  'Checkbox
         TabIndex        =   49
         Top             =   3120
         Width           =   4815
      End
      Begin VB.ComboBox cmb_407Anti 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "frm_mo.frx":21D5
         Left            =   -72480
         List            =   "frm_mo.frx":21E5
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   2640
         Width           =   3495
      End
      Begin VB.ListBox ListG407Gpos 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   -73800
         Style           =   1  'Checkbox
         TabIndex        =   47
         Top             =   840
         Width           =   4815
      End
      Begin VB.ListBox ListG407Gneg 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   -67800
         Style           =   1  'Checkbox
         TabIndex        =   46
         Top             =   840
         Width           =   4935
      End
      Begin VB.CommandButton cmd_Add407Gpos 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Add Gram positive"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton cmd_Add407Gneg 
         BackColor       =   &H0080FF80&
         Caption         =   "Add Gram negative"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -68880
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   1200
         Width           =   975
      End
      Begin VB.ComboBox cmb_407GramPos 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "frm_mo.frx":220E
         Left            =   -72840
         List            =   "frm_mo.frx":221E
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   360
         Width           =   2655
      End
      Begin VB.ComboBox cmb_407GramNeg 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "frm_mo.frx":2247
         Left            =   -65520
         List            =   "frm_mo.frx":2257
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox Txt_407Gpos 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   -73800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   41
         Top             =   2040
         Width           =   4815
      End
      Begin VB.TextBox Txt_407Gneg 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   -67800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Top             =   2040
         Width           =   4935
      End
      Begin VB.ComboBox cmb_ACTIVEYN 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "frm_mo.frx":2280
         Left            =   -69000
         List            =   "frm_mo.frx":228D
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmd_Go 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Go"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68280
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Txt_BacteriaG410 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   -68640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   37
         Top             =   2040
         Width           =   6015
      End
      Begin VB.Label Label14 
         Caption         =   "Virus"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   119
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Bacteria"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68400
         TabIndex        =   118
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Bacteria"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   117
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Virus"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   116
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Yeast"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8400
         TabIndex        =   115
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Parasites"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -66600
         TabIndex        =   114
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Virus"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74640
         TabIndex        =   113
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Bacteria"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70800
         TabIndex        =   112
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Bacteria"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68520
         TabIndex        =   111
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Virus"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74400
         TabIndex        =   110
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Antimicrobial Resistance Genes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   109
         Top             =   2760
         Width           =   2775
      End
      Begin VB.Label Label12 
         Caption         =   "Atypical Bacteria"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68640
         TabIndex        =   108
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Yeast"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68400
         TabIndex        =   107
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label Label17 
         Caption         =   "Antibiotic resistance genes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   106
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Label Label18 
         Caption         =   "Gram positive bacteria"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   105
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label19 
         Caption         =   "Gram-negative bacteria"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -67680
         TabIndex        =   104
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label20 
         Caption         =   "NEW PANEL? "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70080
         TabIndex        =   103
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Label lbl_microbiology 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MOLECULAR MICROBIOLOGY REPORT SCREEN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   345
      Left            =   3120
      TabIndex        =   124
      Top             =   480
      Width           =   7380
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TMH"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3600
      TabIndex        =   123
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   122
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label lbl_final 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Finalized By"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   5760
      TabIndex        =   121
      Top             =   8160
      Width           =   1155
   End
   Begin VB.Label lbl_enterby 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entered  By"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   600
      TabIndex        =   120
      Top             =   8160
      Width           =   1155
   End
End
Attribute VB_Name = "frm_mo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Krsna
'Option Explicit
Public a As Integer
Public req

Dim m
Dim X As Integer
Dim SEL_DMUL_QR As rdoQuery
Dim SEL_DMUL_RS As rdoResultset

Dim SEL_MOVIRMST_QR As rdoQuery
Dim SEL_MOVIRMST_RS As rdoResultset

Dim SEL_MOBACTMST_QR As rdoQuery
Dim SEL_MOBACTMST_RS As rdoResultset

Dim SEL_MOYEASTMST_QR As rdoQuery
Dim SEL_MOYEASTMST_RS As rdoResultset

Dim SEL_MOPARAMST_QR As rdoQuery
Dim SEL_MOPARAMST_RS As rdoResultset

Dim SEL_MOANTIMST_QR As rdoQuery
Dim SEL_MOANTIMST_RS As rdoResultset

Dim SEL_MOATYPMST_QR As rdoQuery
Dim SEL_MOATYPMST_RS As rdoResultset

Dim SEL_408TRN_QR As rdoQuery
Dim SEL_408TRN_RS As rdoResultset

Dim SEL_409TRN_QR As rdoQuery
Dim SEL_409TRN_RS As rdoResultset

Dim SEL_406TRN_QR As rdoQuery
Dim SEL_406TRN_RS As rdoResultset

Dim SEL_410TRN_QR As rdoQuery
Dim SEL_410TRN_RS As rdoResultset

Dim SEL_407TRN_QR As rdoQuery
Dim SEL_407TRN_RS As rdoResultset

Dim INS_408TRN_QR As rdoQuery
Dim INS_409TRN_QR As rdoQuery
Dim INS_406TRN_QR As rdoQuery
Dim INS_410TRN_QR As rdoQuery
Dim INS_407TRN_QR As rdoQuery

Dim UPD_408TRN_QR As rdoQuery
Dim UPD_409TRN_QR As rdoQuery
Dim UPD_406TRN_QR As rdoQuery
Dim UPD_410TRN_QR As rdoQuery
Dim UPD_407TRN_QR As rdoQuery

'***************

Dim req_qr As rdoQuery
Dim req_rs As rdoResultset

Dim sel_Vreq_qr As rdoQuery
Dim sel_Vreq_rs As rdoResultset

Dim NEWFLAG As Boolean
Dim viewflag As Boolean
Dim flagsave As Integer

Dim SELTA101_QR As rdoQuery
Dim SELTA101_RS As rdoResultset

Dim SELPATHDREQUISITN_QR As rdoQuery
Dim selpathdrequisitn_rs As rdoResultset

Dim REFDR_QR As rdoQuery
Dim REFDR_RS As rdoResultset

Dim upd_Trrequisitn_qr As rdoQuery
Dim updrrequisitn1_qr As rdoQuery

Dim UPDRTIME_QR As rdoQuery
Dim UPDRTIMEC_QR As rdoQuery

Dim USRNAME_QR As rdoQuery
Dim USRNAME_RS As rdoResultset

Dim SELDREQSTAFFDQR As rdoQuery
Dim SELDREQSTAFFDRS As rdoResultset

Dim LOGINUSRNAME As String
Dim LOGINTYP As String
Dim LOGINUSRID As String

Dim CURUSRDET_QR As rdoQuery
Dim CURUSRDET_RS As rdoResultset



Private Sub cmb_408BactDet_Click()
If cmb_408BactDet.Text = "DETECTED" Then
    ListBacteriaG408.Enabled = True
    Txt_BacteriaG408.Enabled = True
    cmd_addB408.Enabled = True
Else
    ListBacteriaG408.Enabled = False
    Txt_BacteriaG408.Enabled = False
    cmd_addB408.Enabled = False
    
    Txt_BacteriaG408.Text = cmb_408BactDet.Text
End If
End Sub

Private Sub cmb_408virDet_Click()
If cmb_408virDet.Text = "DETECTED" Then
    ListVirusG408.Enabled = True
    Txt_VirusG408.Enabled = True
    cmd_AddV408.Enabled = True
Else
    ListVirusG408.Enabled = False
    Txt_VirusG408.Enabled = False
    cmd_AddV408.Enabled = False
    
    Txt_VirusG408.Text = Trim(cmb_408virDet.Text)
End If
End Sub

Private Sub cmb_409BactDet_Click()
If cmb_409BactDet.Text = "DETECTED" Then
    ListBacteriaG409.Enabled = True
    Txt_BacteriaG409.Enabled = True
    cmd_addB409.Enabled = True
    Txt_BacteriaG409.Text = ""
Else
    ListBacteriaG409.Enabled = False
    Txt_BacteriaG409.Enabled = False
    cmd_addB409.Enabled = False
    
    Txt_BacteriaG409.Text = cmb_409BactDet.Text
End If

End Sub

Private Sub cmb_406BactDet_Click()
If cmb_406BactDet.Text = "DETECTED" Then
    ListBacteriaG406.Enabled = True
    Txt_BacteriaG406.Enabled = True
    cmd_addB406.Enabled = True
    Txt_BacteriaG406.Text = ""
Else
    ListBacteriaG406.Enabled = False
    Txt_BacteriaG406.Enabled = False
    cmd_addB406.Enabled = False
    
    Txt_BacteriaG406.Text = cmb_406BactDet.Text
End If

End Sub



Private Sub cmb_410BactDet_Click()
If cmb_410BactDet.Text = "DETECTED" Then
    ListBacteriaG410.Enabled = True
    Txt_BacteriaG410.Enabled = True
    cmd_addB410.Enabled = True
    Txt_BacteriaG410.Text = ""
Else
    ListBacteriaG410.Enabled = False
    Txt_BacteriaG410.Enabled = False
    cmd_addB410.Enabled = False
    
    Txt_BacteriaG410.Text = cmb_410BactDet.Text
End If

End Sub

Private Sub cmb_407GramPos_Click()
If cmb_407GramPos.Text = "DETECTED" Then
    ListG407Gpos.Enabled = True
    Txt_407Gpos.Enabled = True
    cmd_Add407Gpos.Enabled = True
    Txt_407Gpos.Text = ""
Else
    ListG407Gpos.Enabled = False
    Txt_407Gpos.Enabled = False
    cmd_Add407Gpos.Enabled = False
    
    Txt_407Gpos.Text = cmb_407GramPos.Text
End If

End Sub

Private Sub cmb_409virDet_Click()
If cmb_409virDet.Text = "DETECTED" Then
    ListVirusG409.Enabled = True
    Txt_VirusG409.Enabled = True
    cmd_AddV409.Enabled = True
    Txt_VirusG409.Text = ""
Else
    ListVirusG409.Enabled = False
    Txt_VirusG409.Enabled = False
    cmd_AddV409.Enabled = False
    
    Txt_VirusG409.Text = Trim(cmb_409virDet.Text)
End If
End Sub

Private Sub cmb_406virDet_Click()
If cmb_406virDet.Text = "DETECTED" Then
    ListVirusG406.Enabled = True
    Txt_VirusG406.Enabled = True
    cmd_AddV406.Enabled = True
    Txt_VirusG406.Text = ""
Else
    ListVirusG406.Enabled = False
    Txt_VirusG406.Enabled = False
    cmd_AddV406.Enabled = False
    
    Txt_VirusG406.Text = Trim(cmb_406virDet.Text)
End If
End Sub

Private Sub cmb_410virDet_Click()
If cmb_410virDet.Text = "DETECTED" Then
    ListVirusG410.Enabled = True
    Txt_VirusG410.Enabled = True
    cmd_AddV410.Enabled = True
    Txt_VirusG410.Text = ""
Else
    ListVirusG410.Enabled = False
    Txt_VirusG410.Enabled = False
    cmd_AddV410.Enabled = False
    
    Txt_VirusG410.Text = Trim(cmb_410virDet.Text)
End If
End Sub

Private Sub cmb_407GramNeg_Click()
If cmb_407GramNeg.Text = "DETECTED" Then
    ListG407Gneg.Enabled = True
    Txt_407Gneg.Enabled = True
    cmd_Add407Gneg.Enabled = True
    Txt_407Gneg.Text = ""
Else
    ListG407Gneg.Enabled = False
    Txt_407Gneg.Enabled = False
    cmd_Add407Gneg.Enabled = False
    
    Txt_407Gneg.Text = cmb_407GramNeg.Text
End If

End Sub



Private Sub cmb_409YeastDet_Click()
If cmb_409YeastDet.Text = "DETECTED" Then
    ListYeastG409.Enabled = True
    Txt_YeastG409.Enabled = True
    cmd_addY409.Enabled = True
    Txt_YeastG409.Text = ""
Else
    ListYeastG409.Enabled = False
    Txt_YeastG409.Enabled = False
    cmd_addY409.Enabled = False
    
    Txt_YeastG409.Text = cmb_409YeastDet.Text
End If
End Sub

Private Sub cmb_406ParaDet_Click()
If cmb_406ParaDet.Text = "DETECTED" Then
    ListParaG406.Enabled = True
    Txt_ParaG406.Enabled = True
    cmd_addP406.Enabled = True
    Txt_ParaG406.Text = ""
Else
    ListParaG406.Enabled = False
    Txt_ParaG406.Enabled = False
    cmd_addP406.Enabled = False
    
    Txt_ParaG406.Text = cmb_406ParaDet.Text
End If
End Sub

Private Sub cmb_410AntiDet_Click()
If cmb_410AntiDet.Text = "DETECTED" Then
    ListAntiG410.Enabled = True
    Txt_AntiG410.Enabled = True
    cmd_AddAnti410.Enabled = True
    Txt_AntiG410.Text = ""
Else
    ListAntiG410.Enabled = False
    Txt_AntiG410.Enabled = False
    cmd_AddAnti410.Enabled = False
    
    Txt_AntiG410.Text = cmb_410AntiDet.Text
End If
End Sub

Private Sub cmb_407Anti_Click()
If cmb_407Anti.Text = "DETECTED" Then
    ListG407Anti.Enabled = True
    Txt_407Anti.Enabled = True
    cmd_Add407Ganti.Enabled = True
    Txt_407Anti.Text = ""
Else
    ListG407Anti.Enabled = False
    Txt_407Anti.Enabled = False
    cmd_Add407Ganti.Enabled = False
    
    Txt_407Anti.Text = cmb_407Anti.Text
End If
End Sub

Private Sub cmb_410AtypicalDet_Click()
If cmb_410AtypicalDet.Text = "DETECTED" Then
    ListAtypicalG410.Enabled = True
    Txt_AtypicalG410.Enabled = True
    cmd_AddAtypical410.Enabled = True
    Txt_AtypicalG410.Text = ""
Else
    ListAtypicalG410.Enabled = False
    Txt_AtypicalG410.Enabled = False
    cmd_AddAtypical410.Enabled = False
    
    Txt_AtypicalG410.Text = cmb_410AtypicalDet.Text
End If
End Sub

Private Sub cmb_407Yeast_Click()
If cmb_407Yeast.Text = "DETECTED" Then
    ListG407Yeast.Enabled = True
    Txt_407Yeast.Enabled = True
    cmd_Add407Gyeast.Enabled = True
    Txt_407Yeast.Text = ""
Else
    ListG407Yeast.Enabled = False
    Txt_407Yeast.Enabled = False
    cmd_Add407Gyeast.Enabled = False
    
    Txt_407Yeast.Text = cmb_407Yeast.Text
End If
End Sub

Private Sub cmb_finalby_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then KeyCode = 0
End Sub

Private Sub cmb_finalby_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Else
KeyAscii = 0
End If
End Sub

Private Sub cmd_addB408_Click()
Txt_BacteriaG408.Text = ""

For X = 0 To ListBacteriaG408.ListCount - 1
    If ListBacteriaG408.Selected(X) = True Then
        Txt_BacteriaG408.Text = Txt_BacteriaG408.Text & ", " & Trim(ListBacteriaG408.List(X))
    End If
Next

If Trim(Txt_BacteriaG408.Text) <> "" Then
    Txt_BacteriaG408.Text = Mid(Trim(Txt_BacteriaG408.Text), 2, Len(Trim(Txt_BacteriaG408.Text)) - 1)
End If
End Sub

Private Sub cmd_addB409_Click()
Txt_BacteriaG409.Text = ""

For X = 0 To ListBacteriaG409.ListCount - 1
    If ListBacteriaG409.Selected(X) = True Then
        Txt_BacteriaG409.Text = Txt_BacteriaG409.Text & ", " & Trim(ListBacteriaG409.List(X))
    End If
Next

If Trim(Txt_BacteriaG409.Text) <> "" Then
    Txt_BacteriaG409.Text = Mid(Trim(Txt_BacteriaG409.Text), 2, Len(Trim(Txt_BacteriaG409.Text)) - 1)
End If
End Sub

Private Sub cmd_addB406_Click()
Txt_BacteriaG406.Text = ""

For X = 0 To ListBacteriaG406.ListCount - 1
    If ListBacteriaG406.Selected(X) = True Then
        Txt_BacteriaG406.Text = Txt_BacteriaG406.Text & ", " & Trim(ListBacteriaG406.List(X))
    End If
Next

If Trim(Txt_BacteriaG406.Text) <> "" Then
    Txt_BacteriaG406.Text = Mid(Trim(Txt_BacteriaG406.Text), 2, Len(Trim(Txt_BacteriaG406.Text)) - 1)
End If
End Sub

Private Sub cmd_addB410_Click()
Txt_BacteriaG410.Text = ""

For X = 0 To ListBacteriaG410.ListCount - 1
    If ListBacteriaG410.Selected(X) = True Then
        Txt_BacteriaG410.Text = Txt_BacteriaG410.Text & ", " & Trim(ListBacteriaG410.List(X))
    End If
Next

If Trim(Txt_BacteriaG410.Text) <> "" Then
    Txt_BacteriaG410.Text = Mid(Trim(Txt_BacteriaG410.Text), 2, Len(Trim(Txt_BacteriaG410.Text)) - 1)
End If
End Sub


Private Sub cmd_Add407Gpos_Click()
Txt_407Gpos.Text = ""

For X = 0 To ListG407Gpos.ListCount - 1
    If ListG407Gpos.Selected(X) = True Then
        Txt_407Gpos.Text = Txt_407Gpos.Text & ", " & Trim(ListG407Gpos.List(X))
    End If
Next

If Trim(Txt_407Gpos.Text) <> "" Then
    Txt_407Gpos.Text = Mid(Trim(Txt_407Gpos.Text), 2, Len(Trim(Txt_407Gpos.Text)) - 1)
End If
End Sub

Private Sub cmd_AddV408_Click()
Txt_VirusG408.Text = ""

For X = 0 To ListVirusG408.ListCount - 1
    If ListVirusG408.Selected(X) = True Then
        Txt_VirusG408.Text = Txt_VirusG408.Text & ", " & Trim(ListVirusG408.List(X))
    End If
Next

If Trim(Txt_VirusG408.Text) <> "" Then
    Txt_VirusG408.Text = Mid(Trim(Txt_VirusG408.Text), 2, Len(Trim(Txt_VirusG408.Text)) - 1)
End If
End Sub

Private Sub cmd_AddV409_Click()
Txt_VirusG409.Text = ""

For X = 0 To ListVirusG409.ListCount - 1
    If ListVirusG409.Selected(X) = True Then
        Txt_VirusG409.Text = Txt_VirusG409.Text & ", " & Trim(ListVirusG409.List(X))
    End If
Next

If Trim(Txt_VirusG409.Text) <> "" Then
    Txt_VirusG409.Text = Mid(Trim(Txt_VirusG409.Text), 2, Len(Trim(Txt_VirusG409.Text)) - 1)
End If
End Sub

Private Sub cmd_AddV406_Click()
Txt_VirusG406.Text = ""

For X = 0 To ListVirusG406.ListCount - 1
    If ListVirusG406.Selected(X) = True Then
        Txt_VirusG406.Text = Txt_VirusG406.Text & ", " & Trim(ListVirusG406.List(X))
    End If
Next

If Trim(Txt_VirusG406.Text) <> "" Then
    Txt_VirusG406.Text = Mid(Trim(Txt_VirusG406.Text), 2, Len(Trim(Txt_VirusG406.Text)) - 1)
End If
End Sub

Private Sub cmd_AddV410_Click()
Txt_VirusG410.Text = ""

For X = 0 To ListVirusG410.ListCount - 1
    If ListVirusG410.Selected(X) = True Then
        Txt_VirusG410.Text = Txt_VirusG410.Text & ", " & Trim(ListVirusG410.List(X))
    End If
Next

If Trim(Txt_VirusG410.Text) <> "" Then
    Txt_VirusG410.Text = Mid(Trim(Txt_VirusG410.Text), 2, Len(Trim(Txt_VirusG410.Text)) - 1)
End If
End Sub

Private Sub cmd_Add407Gneg_Click()
Txt_407Gneg.Text = ""

For X = 0 To ListG407Gneg.ListCount - 1
    If ListG407Gneg.Selected(X) = True Then
        Txt_407Gneg.Text = Txt_407Gneg.Text & ", " & Trim(ListG407Gneg.List(X))
    End If
Next

If Trim(Txt_407Gneg.Text) <> "" Then
    Txt_407Gneg.Text = Mid(Trim(Txt_407Gneg.Text), 2, Len(Trim(Txt_407Gneg.Text)) - 1)
End If
End Sub

Private Sub cmd_addY409_Click()
Txt_YeastG409.Text = ""

For X = 0 To ListYeastG409.ListCount - 1
    If ListYeastG409.Selected(X) = True Then
        Txt_YeastG409.Text = Txt_YeastG409.Text & ", " & Trim(ListYeastG409.List(X))
    End If
Next

If Trim(Txt_YeastG409.Text) <> "" Then
    Txt_YeastG409.Text = Mid(Trim(Txt_YeastG409.Text), 2, Len(Trim(Txt_YeastG409.Text)) - 1)
End If
End Sub

Private Sub cmd_AddAnti410_Click()
Txt_AntiG410.Text = ""

For X = 0 To ListAntiG410.ListCount - 1
    If ListAntiG410.Selected(X) = True Then
        Txt_AntiG410.Text = Txt_AntiG410.Text & ", " & Trim(ListAntiG410.List(X))
    End If
Next

If Trim(Txt_AntiG410.Text) <> "" Then
    Txt_AntiG410.Text = Mid(Trim(Txt_AntiG410.Text), 2, Len(Trim(Txt_AntiG410.Text)) - 1)
End If
End Sub

Private Sub cmd_Add407Ganti_Click()
Txt_407Anti.Text = ""

For X = 0 To ListG407Anti.ListCount - 1
    If ListG407Anti.Selected(X) = True Then
        Txt_407Anti.Text = Txt_407Anti.Text & ", " & Trim(ListG407Anti.List(X))
    End If
Next

If Trim(Txt_407Anti.Text) <> "" Then
    Txt_407Anti.Text = Mid(Trim(Txt_407Anti.Text), 2, Len(Trim(Txt_407Anti.Text)) - 1)
End If
End Sub

Private Sub cmd_AddAtypical410_Click()
Txt_AtypicalG410.Text = ""

For X = 0 To ListAtypicalG410.ListCount - 1
    If ListAtypicalG410.Selected(X) = True Then
        Txt_AtypicalG410.Text = Txt_AtypicalG410.Text & ", " & Trim(ListAtypicalG410.List(X))
    End If
Next

If Trim(Txt_AtypicalG410.Text) <> "" Then
    Txt_AtypicalG410.Text = Mid(Trim(Txt_AtypicalG410.Text), 2, Len(Trim(Txt_AtypicalG410.Text)) - 1)
End If
End Sub


Private Sub cmd_Add407Gyeast_Click()
Txt_407Yeast.Text = ""

For X = 0 To ListG407Yeast.ListCount - 1
    If ListG407Yeast.Selected(X) = True Then
        Txt_407Yeast.Text = Txt_407Yeast.Text & ", " & Trim(ListG407Yeast.List(X))
    End If
Next

If Trim(Txt_407Yeast.Text) <> "" Then
    Txt_407Yeast.Text = Mid(Trim(Txt_407Yeast.Text), 2, Len(Trim(Txt_407Yeast.Text)) - 1)
End If
End Sub


Private Sub cmd_addP406_Click()
Txt_ParaG406.Text = ""

For X = 0 To ListParaG406.ListCount - 1
    If ListParaG406.Selected(X) = True Then
        Txt_ParaG406.Text = Txt_ParaG406.Text & ", " & Trim(ListParaG406.List(X))
    End If
Next

If Trim(Txt_ParaG406.Text) <> "" Then
    Txt_ParaG406.Text = Mid(Trim(Txt_ParaG406.Text), 2, Len(Trim(Txt_ParaG406.Text)) - 1)
End If
End Sub

Private Sub cmd_commit_Click()
str = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=CURRENT DATE, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
Set updrrequisitn1_qr = rdb.CreateQuery("", str)

str = "UPDATE RISLIB.RTIMELOG SET COMREPDATE=CURRENT DATE,COMREPTIME=CURRENT TIME WHERE REQNNO=?"
Set UPDRTIMEC_QR = rdb.CreateQuery("", str)

Dim i As Integer
If Len(Trim(cmb_finalby)) = 0 Then
    m = MsgBox("The Finalized by field is EMPTY.", vbCritical)
    Exit Sub
End If
i = MsgBox(" Do you want to Commit the Record?", vbYesNo, vbInformation)


If i = vbYes Then
    On Error GoTo MSG:

    rdb.BeginTrans
    updrrequisitn1_qr(0) = UCase(cmb_finalby.Text)
    
    updrrequisitn1_qr(1) = "F"
    updrrequisitn1_qr(2) = "Y"
    updrrequisitn1_qr(3) = Trim(cmb_reqno.Text)
    updrrequisitn1_qr.Execute
    
    
    UPDRTIMEC_QR(0) = Trim(cmb_reqno.Text)
    UPDRTIMEC_QR.Execute


    '***** SMS *****
   
   ' sendRMSG Trim(cmb_reqno.Text), MaskEdBox1.FormattedText, "F", WIN_CLIENT.LocalIP  'commented by varinder
    
    '*******************************

    
    MsgBox "Record is Committed Successfully."
End If

rdb.CommitTrans
BTNENABL1 ("INI")
cmd_print.Enabled = True
Exit Sub
MSG:
rdb.RollbackTrans
MsgBox Err.Description
End Sub

Private Sub cmd_Go_Click()

    ListG407Gpos.clear
    
    str = "select * FROM DISLIB.MOGRAMPMST WHERE SER_CODE='G407' AND ACTIVEYN=? ORDER BY BACTERIA"
    Set SEL_MOBACTMST_QR = rdb.CreateQuery("", str)
    
    SEL_MOBACTMST_QR(0) = Trim(cmb_ACTIVEYN.Text)
    Set SEL_MOBACTMST_RS = SEL_MOBACTMST_QR.OpenResultset(1, 3)
    While Not SEL_MOBACTMST_RS.EOF
        ListG407Gpos.AddItem Trim(SEL_MOBACTMST_RS("BACTERIA"))
        SEL_MOBACTMST_RS.MoveNext
    Wend
    
    '*****
    ListG407Gneg.clear
    
    str = "select * FROM DISLIB.MOGRAMNMST WHERE SER_CODE='G407' AND ACTIVEYN=? ORDER BY BACTERIA"
    Set SEL_MOBACTMST_QR = rdb.CreateQuery("", str)
    
    SEL_MOBACTMST_QR(0) = Trim(cmb_ACTIVEYN.Text)
    Set SEL_MOBACTMST_RS = SEL_MOBACTMST_QR.OpenResultset(1, 3)
    While Not SEL_MOBACTMST_RS.EOF
        ListG407Gneg.AddItem Trim(SEL_MOBACTMST_RS("BACTERIA"))
        SEL_MOBACTMST_RS.MoveNext
    Wend
    
    '*****
    ListG407Anti.clear
    
    str = "select * FROM DISLIB.MORSGENMST WHERE SER_CODE='G407' AND ACTIVEYN=? ORDER BY RESGENES"
    Set SEL_MOANTIMST_QR = rdb.CreateQuery("", str)
    
    SEL_MOANTIMST_QR(0) = Trim(cmb_ACTIVEYN.Text)
    Set SEL_MOANTIMST_RS = SEL_MOANTIMST_QR.OpenResultset(1, 3)
    While Not SEL_MOANTIMST_RS.EOF
        ListG407Anti.AddItem Trim(SEL_MOANTIMST_RS("RESGENES"))
        SEL_MOANTIMST_RS.MoveNext
    Wend
  '*****
  
    ListG407Yeast.clear
    str = "select * FROM DISLIB.MOYEASTMST WHERE SER_CODE='G407' AND ACTIVEYN=? ORDER BY YEAST"
    Set SEL_MOYEASTMST_QR = rdb.CreateQuery("", str)
    
    SEL_MOYEASTMST_QR(0) = Trim(cmb_ACTIVEYN.Text)
    Set SEL_MOYEASTMST_RS = SEL_MOYEASTMST_QR.OpenResultset(1, 3)
    While Not SEL_MOYEASTMST_RS.EOF
        ListG407Yeast.AddItem Trim(SEL_MOYEASTMST_RS("YEAST"))
        SEL_MOYEASTMST_RS.MoveNext
    Wend
End Sub

Private Sub Command1_Click()
'----shweta
    str = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USR_ID='A" & usr_name & "' "
    Set REPWTRN_QR = rdb.CreateQuery("", str)
   Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)
    If REPWTRNRS.RowCount > 0 Then
    cmb_finalby.Text = REPWTRNRS("USERNAME")
    End If
    '----
' Label27.Caption = ""
 
'Dim b As Integer
'b = InStr(1, cmb_reqno, "$")
'mod_reqno = req_chglen(Mid(Trim(cmb_reqno), 1, b - 1)) 'put change value to this var so that it used later on
'
'UPDRREQ_QR(0) = "T"
'UPDRREQ_QR(1) = Null
'If Report_microbiology.Text2.Text <> "" Then
'If Report_microbiology.Text2.Text = "2" Then
''cmb_reqno.Text = req
'UPDRREQ_QR(2) = mod_reqno
'Else
'UPDRREQ_QR(2) = Report_microbiology.Text2.Text
' End If
'  End If
  
'UPDRREQ_QR(2) = mod_reqno
'UPDRREQ_QR.Execute
'
'cmb_reqno_Click
cbccommit.show
Me.Hide
cbccommit.cmd_showrpt_Click
End Sub

Private Sub Form_Load()
SSTab_MO.Visible = False
NEWFLAG = False
viewflag = False

BTNENABL1 ("INI")
FLDENABL1 ("INI")

MaskEdBox1.Enabled = False
datetime

str = "SELECT * FROM DISLIB.DISLNKUSR WHERE USR_ID = ? "
Set USRNAME_QR = rdb.CreateQuery("", str)

USRNAME_QR(0) = UID
Set USRNAME_RS = USRNAME_QR.OpenResultset(1, 1)

If Not USRNAME_RS.EOF Then
    LOGINUSRNAME = Trim(USRNAME_RS("USERNAME"))
    LOGINTYP = Trim(USRNAME_RS("USRTYPE"))
    LOGINUSRID = Trim(USRNAME_RS("usr_id"))
    cmb_enteryby.Text = (LOGINUSRNAME)
    cmb_enteryby.AddItem LOGINUSRNAME
    cmb_enteryby.Text = cmb_enteryby.List(0)
    cmb_enteryby.Enabled = False
Else
    MsgBox "The Security Details of this User is not available!!", vbInformation, "DIS"
    Exit Sub
End If

str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO = ? "
Set SEL_DMUL_QR = rdb.CreateQuery("", str)

str = "select * from DISLIB.Dislnkusr where username = ?"
Set CURUSRDET_QR = rdb.CreateQuery("", str)

'******* Shraddha 10/03/2014 **********
Dim REPWTRNF_QR As rdoQuery
Dim REPWTRNRSF As rdoResultset

str = "SELECT DISTINCT USERNAME,USRTYPE FROM DISLIB.DISLNKUSR A,EMPLIB.ELEMPINFO B WHERE SUBSTR(A.USR_ID,2,7)=B.EMPCODE AND (A.USR_ID LIKE 'A1%') AND (B.DEPTCODE='ACEB' OR B.DEPTCODE='GZZ') AND B.CESSYN='4'"
'str = "SELECT  USERNAME,USRTYPE FROM DISLIB.DISLNKUSR A, EMPLIB.ELEMPINFO B WHERE SUBSTR(A.USR_ID,2,7)=B.EMPCODE AND (A.USR_ID LIKE 'A1%' OR A.USR_ID LIKE 'A2%' OR A.USR_ID LIKE 'A3%') AND B.CESSYN='4' AND MODMR='7' ORDER BY USERNAME"
Set REPWTRNF_QR = rdb.CreateQuery("REPWTRNDDF_STR", str)
Set REPWTRNRSF = REPWTRNF_QR.OpenResultset(1, 1)
If REPWTRNRSF.RowCount > 0 Then

    REPWTRNRSF.MoveFirst
    While Not REPWTRNRSF.EOF
        If REPWTRNRSF("USRTYPE") >= 8 Then
            cmb_finalby.AddItem Trim(REPWTRNRSF("USERNAME"))
        End If
        REPWTRNRSF.MoveNext
    Wend
End If

cmb_408virDet.Enabled = False
cmb_408BactDet.Enabled = False

cmb_409virDet.Enabled = False
cmb_409BactDet.Enabled = False
cmb_409YeastDet.Enabled = False

cmb_406virDet.Enabled = False
cmb_406BactDet.Enabled = False
cmb_406ParaDet.Enabled = False

cmb_410virDet.Enabled = False
cmb_410BactDet.Enabled = False
cmb_410AntiDet.Enabled = False
cmb_410AtypicalDet.Enabled = False

cmb_407GramPos.Enabled = False
cmb_407GramNeg.Enabled = False
cmb_407Anti.Enabled = False
cmb_407Yeast.Enabled = False

ListVirusG408.Enabled = False
ListVirusG409.Enabled = False
ListVirusG406.Enabled = False
ListVirusG410.Enabled = False


ListBacteriaG408.Enabled = False
ListBacteriaG409.Enabled = False
ListBacteriaG406.Enabled = False
ListBacteriaG410.Enabled = False

ListYeastG409.Enabled = False
ListParaG406.Enabled = False
ListAntiG410.Enabled = False
ListAtypicalG410.Enabled = False


ListG407Gpos.Enabled = False
ListG407Gneg.Enabled = False
ListG407Anti.Enabled = False
ListG407Yeast.Enabled = False

Txt_VirusG408.Enabled = False
Txt_VirusG409.Enabled = False
Txt_VirusG406.Enabled = False
Txt_VirusG410.Enabled = False

Txt_BacteriaG408.Enabled = False
Txt_BacteriaG409.Enabled = False
Txt_BacteriaG406.Enabled = False
Txt_BacteriaG410.Enabled = False

Txt_YeastG409.Enabled = False
Txt_ParaG406.Enabled = False
Txt_AntiG410.Enabled = False
Txt_AtypicalG410.Enabled = False

Txt_407Gpos.Enabled = False
Txt_407Gneg.Enabled = False
Txt_407Anti.Enabled = False
Txt_407Yeast.Enabled = False

txt_MOremarks.Text = ""
Txt_VirusG408.Text = ""
Txt_VirusG409.Text = ""
Txt_VirusG406.Text = ""
Txt_VirusG410.Text = ""

Txt_BacteriaG408.Text = ""
Txt_BacteriaG409.Text = ""
Txt_BacteriaG406.Text = ""
Txt_BacteriaG410.Text = ""

Txt_YeastG409.Text = ""
Txt_ParaG406.Text = ""
Txt_AntiG410.Text = ""
Txt_AtypicalG410.Text = ""

Txt_407Gpos.Text = ""
Txt_407Gneg.Text = ""
Txt_407Anti.Text = ""
Txt_407Yeast.Text = ""

txt_MOremarks.Text = ""

cmd_AddV408.Enabled = False
cmd_AddV409.Enabled = False
cmd_AddV406.Enabled = False
cmd_AddV410.Enabled = False

cmd_addB408.Enabled = False
cmd_addB409.Enabled = False
cmd_addB406.Enabled = False
cmd_addB410.Enabled = False

cmd_addY409.Enabled = False
cmd_addP406.Enabled = False
cmd_AddAnti410.Enabled = False
cmd_AddAtypical410.Enabled = False

cmd_Add407Gpos.Enabled = False
cmd_Add407Gneg.Enabled = False
cmd_Add407Ganti.Enabled = False
cmd_Add407Gyeast.Enabled = False

SSTab_MO.Visible = True
SSTab_MO.TabEnabled(0) = False
SSTab_MO.TabEnabled(1) = False
SSTab_MO.TabEnabled(2) = False
SSTab_MO.TabEnabled(3) = False
SSTab_MO.TabEnabled(4) = False

'SSTab_MO.TabVisible(0) = False
'SSTab_MO.TabVisible(1) = False
'SSTab_MO.TabVisible(2) = False
'SSTab_MO.TabVisible(3) = False
'SSTab_MO.TabVisible(4) = False
End Sub


Public Sub cmd_new_Click()
BTNENABL1 ("INI")
FLDENABL1 ("INI")
SSTab_MO.Visible = False
txt_MOremarks.Text = ""
Txt_VirusG408.Text = ""
Txt_VirusG409.Text = ""
Txt_VirusG406.Text = ""
Txt_VirusG410.Text = ""

Txt_BacteriaG408.Text = ""
Txt_BacteriaG409.Text = ""
Txt_BacteriaG406.Text = ""
Txt_BacteriaG410.Text = ""

Txt_YeastG409.Text = ""
Txt_ParaG406.Text = ""
Txt_AntiG410.Text = ""
Txt_AtypicalG410.Text = ""

Txt_407Gpos.Text = ""
Txt_407Gneg.Text = ""
Txt_407Anti.Text = ""
Txt_407Yeast.Text = ""

txt_MOremarks.Text = ""

BTNENABL1 ("NEW")
FLDENABL1 ("NEW")

viewflag = False
NEWFLAG = True
flagsave = 1
SSTab_MO.Visible = False
End Sub

Public Sub cmd_view_Click()
BTNENABL1 ("INI")
FLDENABL1 ("INI")
SSTab_MO.Visible = False
txt_MOremarks.Text = ""
Txt_VirusG408.Text = ""
Txt_VirusG409.Text = ""
Txt_VirusG406.Text = ""
Txt_VirusG410.Text = ""

Txt_BacteriaG408.Text = ""
Txt_BacteriaG409.Text = ""
Txt_BacteriaG406.Text = ""
Txt_BacteriaG410.Text = ""

Txt_YeastG409.Text = ""
Txt_ParaG406.Text = ""
Txt_AntiG410.Text = ""
Txt_AtypicalG410.Text = ""

Txt_407Gpos.Text = ""
Txt_407Gneg.Text = ""
Txt_407Anti.Text = ""
Txt_407Yeast.Text = ""

txt_MOremarks.Text = ""

BTNENABL1 ("VIEW")
FLDENABL1 ("VIEW")

NEWFLAG = False
viewflag = True
SSTab_MO.Visible = False
End Sub

Private Sub cmd_modify_Click()

BTNENABL1 ("NEW")
FLDENABL1 ("MODIFY")

flagsave = 2

Dim SEL_REQNSTAT_QR As rdoQuery
Dim SEL_REQNSTAT_RS As rdoResultset

str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
Set SEL_REQNSTAT_QR = rdb.CreateQuery("", str)

SEL_REQNSTAT_QR(0) = Trim(cmb_reqno.Text)
Set SEL_REQNSTAT_RS = SEL_REQNSTAT_QR.OpenResultset(1, 3)



End Sub

Private Sub cmd_save_Click()

If Len(Trim(cmb_reqno.Text)) = 0 Then
    MsgBox "Please select REQUISITION NUMBER"
    Exit Sub
End If

If SSTab_MO.TabEnabled(0) = True Then

    If cmb_408virDet.Text = "DETECTED" Or cmb_408virDet.Text = "" Then
    
        If Trim(Txt_VirusG408.Text) = "" Then
            MsgBox "Please select Virus"
            Exit Sub
        End If
    End If
    
    If cmb_408BactDet.Text = "DETECTED" Or cmb_408BactDet.Text = "" Then
        
        If Trim(Txt_BacteriaG408.Text) = "" Then
            MsgBox "Please select Bacteria"
            Exit Sub
        End If

    End If

End If

If SSTab_MO.TabEnabled(1) = True Then

    
    
    If cmb_409virDet.Text = "DETECTED" Or cmb_409virDet.Text = "" Then
    
        If Trim(Txt_VirusG409.Text) = "" Then
            MsgBox "Please select Virus"
            Exit Sub
        End If
    End If
    
    If cmb_409BactDet.Text = "DETECTED" Or cmb_409BactDet.Text = "" Then
        
        If Trim(Txt_BacteriaG409.Text) = "" Then
            MsgBox "Please select Bacteria"
            Exit Sub
        End If

    End If
    
    If cmb_409YeastDet.Text = "DETECTED" Or cmb_409YeastDet.Text = "" Then
        
        If Trim(Txt_YeastG409.Text) = "" Then
            MsgBox "Please select yeast"
            Exit Sub
        End If

    End If

End If

If SSTab_MO.TabEnabled(2) = True Then
    
    If cmb_406virDet.Text = "DETECTED" Or cmb_406virDet.Text = "" Then
    
        If Trim(Txt_VirusG406.Text) = "" Then
            MsgBox "Please select Virus"
            Exit Sub
        End If
    End If
    
    If cmb_406BactDet.Text = "DETECTED" Or cmb_406BactDet.Text = "" Then
        
        If Trim(Txt_BacteriaG406.Text) = "" Then
            MsgBox "Please select Bacteria"
            Exit Sub
        End If

    End If
    
    If cmb_406ParaDet.Text = "DETECTED" Or cmb_406ParaDet.Text = "" Then
        
        If Trim(Txt_ParaG406.Text) = "" Then
            MsgBox "Please select Parasites"
            Exit Sub
        End If

    End If

End If

If SSTab_MO.TabEnabled(3) = True Then
    
    If cmb_410virDet.Text = "DETECTED" Or cmb_410virDet.Text = "" Then
    
        If Trim(Txt_VirusG410.Text) = "" Then
            MsgBox "Please select Virus"
            Exit Sub
        End If
    End If
    
    If cmb_410BactDet.Text = "DETECTED" Or cmb_410BactDet.Text = "" Then
        
        If Trim(Txt_BacteriaG410.Text) = "" Then
            MsgBox "Please select Bacteria"
            Exit Sub
        End If

    End If
    
    If cmb_410AntiDet.Text = "DETECTED" Or cmb_410AntiDet.Text = "" Then
        
        If Trim(Txt_AntiG410.Text) = "" Then
            MsgBox "Please select Antimicrobial Resistance Genes"
            Exit Sub
        End If

    End If
    
    If cmb_410AtypicalDet.Text = "DETECTED" Or cmb_410AtypicalDet.Text = "" Then
        
        If Trim(Txt_AtypicalG410.Text) = "" Then
            MsgBox "Please select Atypical Bacteria"
            Exit Sub
        End If

    End If

End If

If SSTab_MO.TabEnabled(4) = True Then

    If cmb_ACTIVEYN.Text = "" Then
        MsgBox "Please select whether panel is new or not"
        Exit Sub
    End If
    
    If cmb_407GramPos.Text = "DETECTED" Or cmb_407GramPos.Text = "" Then
    
        If Trim(Txt_407Gpos.Text) = "" Then
            MsgBox "Please select Gram positive bacteria"
            Exit Sub
        End If
    End If
    
    If cmb_407GramNeg.Text = "DETECTED" Or cmb_407GramNeg.Text = "" Then
        
        If Trim(Txt_407Gneg.Text) = "" Then
            MsgBox "Please select Gram negative bacteria"
            Exit Sub
        End If

    End If
    
    If cmb_407Anti.Text = "DETECTED" Or cmb_407Anti.Text = "" Then
        
        If Trim(Txt_407Anti.Text) = "" Then
            MsgBox "Please select Antibiotic genes"
            Exit Sub
        End If

    End If
    
    If cmb_407Yeast.Text = "DETECTED" Or cmb_407Yeast.Text = "" Then
        
        If Trim(Txt_407Yeast.Text) = "" Then
            MsgBox "Please select Yeast"
            Exit Sub
        End If

    End If

End If

str = "insert into DISLIB.G408TRN(REQNO,VIRUS,BACTERIA,REMARKS)values(?,?,?,?)"
Set INS_408TRN_QR = rdb.CreateQuery("", str)

str = "insert into DISLIB.G409TRN(REQNO,VIRUS,BACTERIA,YEAST,REMARKS)values(?,?,?,?,?)"
Set INS_409TRN_QR = rdb.CreateQuery("", str)

str = "insert into DISLIB.G406TRN(REQNO,VIRUS,BACTERIA,PARA,REMARKS)values(?,?,?,?,?)"
Set INS_406TRN_QR = rdb.CreateQuery("", str)

str = "insert into DISLIB.G410TRN(REQNO,VIRUS,BACTERIA,ANTI, ATYPICAL,REMARKS)values(?,?,?,?,?,?)"
Set INS_410TRN_QR = rdb.CreateQuery("", str)

str = "insert into DISLIB.G407TRN(REQNO,GPOS,GNEG,ANTI, YEAST,REMARKS,ACTIVEYN)values(?,?,?,?,?,?,?)"
Set INS_407TRN_QR = rdb.CreateQuery("", str)

str = "update dislib.G408TRN set VIRUS=? ,BACTERIA=?,REMARKS=? where reqno=? "
Set UPD_408TRN_QR = rdb.CreateQuery("", str)

str = "update dislib.G409TRN set VIRUS=? ,BACTERIA=?,YEAST=?,REMARKS=? where reqno=? "
Set UPD_409TRN_QR = rdb.CreateQuery("", str)

str = "update dislib.G406TRN set VIRUS=? ,BACTERIA=?,PARA=?,REMARKS=? where reqno=? "
Set UPD_406TRN_QR = rdb.CreateQuery("", str)

str = "update dislib.G410TRN set VIRUS=? ,BACTERIA=?,ANTI=?,ATYPICAL=?,REMARKS=? where reqno=? "
Set UPD_410TRN_QR = rdb.CreateQuery("", str)

str = "update dislib.G407TRN set GPOS=? ,GNEG=?,ANTI=?,YEAST=?,REMARKS=?,ACTIVEYN=? where reqno=? "
Set UPD_407TRN_QR = rdb.CreateQuery("", str)

str = "update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=CURRENT DATE,PREPORTBY=?,REQNSTAT=? WHERE REQNNO=?"
Set upd_Trrequisitn_qr = rdb.CreateQuery("", str)

str = "UPDATE RISLIB.RTIMELOG SET SAVREPDATE=CURRENT DATE,SAVREPTIME=CURRENT TIME WHERE REQNNO=?"
Set UPDRTIME_QR = rdb.CreateQuery("", str)



assign
End Sub

Public Sub cmd_cancel_Click()
BTNENABL1 ("INI")
FLDENABL1 ("INI")
SSTab_MO.Visible = False
End Sub

Private Sub cmd_back_Click()
BTNENABL1 ("INI")
FLDENABL1 ("INI")
Me.Hide
DIS_MAIN.show
End Sub

Private Sub cmd_exit_Click()
End
End Sub

Private Sub cmd_print_Click()

Dim strLink As String
If Left(Trim(cmb_reqno.Text), 3) = "GZZ" Then
    strLink = "http://intranet.tmc.gov.in/emrreports/allreports.aspx?reqno=" & Trim(cmb_reqno.Text) & "~print"
    Call Shell("C:\Program Files\Internet Explorer\iexplore " & strLink, vbMaximizedFocus)
    Exit Sub
End If
End Sub


Public Sub cmb_reqno_Click()

str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO =? "
Set SELPATHDREQUISITN_QR = rdb.CreateQuery("", str)

Set SELTA101_QR = rdb.CreateQuery("", "select * from PATOOLSLIB.ta101 where case_no =?")

str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=? AND EMPCODE=?"
Set SELDREQSTAFFDQR = rdb.CreateQuery("", str)


SELTA101_QR(0) = MaskEdBox1.FormattedText
Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)

If Opt_patient.Value = True Then
    txt_sex.Text = SELTA101_RS("sex")
    txt_name.Text = Trim(SELTA101_RS("first_name")) & " " & Trim(SELTA101_RS("mid_name")) & " " & Trim(SELTA101_RS("last_name"))
    txt_age.Text = DateDiff("YYYY", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
    txt_category.Text = SELTA101_RS("pat_catg")
    txt_DMG.Text = SELTA101_RS("UD_NAME")
ElseIf opt_staff.Value = True Then
    SELDREQSTAFFDQR(0) = Trim(cmb_reqno)
    SELDREQSTAFFDQR(1) = Trim(MaskEdBox1.FormattedText)
    Set SELDREQSTAFFDRS = SELDREQSTAFFDQR.OpenResultset(1, 3)
            
    If Not SELDREQSTAFFDRS.EOF > 0 Then
        txt_sex.Text = SELDREQSTAFFDRS("sex")
        txt_name.Text = SELDREQSTAFFDRS("PNAME")
        txt_age = SELDREQSTAFFDRS("AGE")
        txt_category.Text = SELDREQSTAFFDRS("RELATION")
    End If
End If

SELPATHDREQUISITN_QR(0) = Trim(cmb_reqno.Text)
Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
If Not selpathdrequisitn_rs.EOF Then
    txt_pathno.Text = Trim(selpathdrequisitn_rs("PATHNO"))
    If IsNull((selpathdrequisitn_rs("NMETERIALS"))) Then
     txt_nmaterial.Text = ""
    Else
    
    txt_nmaterial.Text = Trim(selpathdrequisitn_rs("NMETERIALS"))
    End If
End If

SSTab_MO.Visible = True
SSTab_MO.TabEnabled(0) = False
SSTab_MO.TabEnabled(1) = False
SSTab_MO.TabEnabled(2) = False
SSTab_MO.TabEnabled(3) = False
SSTab_MO.TabEnabled(4) = False

'SSTab_MO.TabVisible(0) = False
'SSTab_MO.TabVisible(1) = False
'SSTab_MO.TabVisible(2) = False
'SSTab_MO.TabVisible(3) = False
'SSTab_MO.TabVisible(4) = False
SEL_DMUL_QR(0) = Trim(cmb_reqno.Text)
Set SEL_DMUL_RS = SEL_DMUL_QR.OpenResultset(1, 3)


If NEWFLAG = True Then

If SEL_DMUL_RS("PRO_CODE") = "G408" Then


    cmb_408virDet.Enabled = True
    cmb_408BactDet.Enabled = True

    SSTab_MO.TabEnabled(0) = True
    SSTab_MO.TabVisible(0) = True
    
    str = "select * FROM DISLIB.MOVIRUSMST WHERE SER_CODE='G408' ORDER BY VIRUS"
    Set SEL_MOVIRMST_QR = rdb.CreateQuery("", str)
    Set SEL_MOVIRMST_RS = SEL_MOVIRMST_QR.OpenResultset(1, 3)
    While Not SEL_MOVIRMST_RS.EOF
        ListVirusG408.AddItem Trim(SEL_MOVIRMST_RS("VIRUS"))
        SEL_MOVIRMST_RS.MoveNext
    Wend
    
    '*****
    
    str = "select * FROM DISLIB.MOBACTEMST WHERE SER_CODE='G408' ORDER BY BACTERIA"
    Set SEL_MOBACTMST_QR = rdb.CreateQuery("", str)
    Set SEL_MOBACTMST_RS = SEL_MOBACTMST_QR.OpenResultset(1, 3)
    While Not SEL_MOBACTMST_RS.EOF
        ListBacteriaG408.AddItem Trim(SEL_MOBACTMST_RS("BACTERIA"))
        SEL_MOBACTMST_RS.MoveNext
    Wend
    
   SSTab_MO.Tab = 0
    
ElseIf SEL_DMUL_RS("PRO_CODE") = "G409" Then

    SSTab_MO.TabEnabled(1) = True
    SSTab_MO.TabVisible(1) = True
    cmb_409virDet.Enabled = True
    cmb_409BactDet.Enabled = True
    cmb_409YeastDet.Enabled = True

    str = "select * FROM DISLIB.MOVIRUSMST WHERE SER_CODE='G409' ORDER BY VIRUS"
    Set SEL_MOVIRMST_QR = rdb.CreateQuery("", str)
    Set SEL_MOVIRMST_RS = SEL_MOVIRMST_QR.OpenResultset(1, 3)
    While Not SEL_MOVIRMST_RS.EOF
        ListVirusG409.AddItem Trim(SEL_MOVIRMST_RS("VIRUS"))
        SEL_MOVIRMST_RS.MoveNext
    Wend
    
    '*****
    
    str = "select * FROM DISLIB.MOBACTEMST WHERE SER_CODE='G409' ORDER BY BACTERIA"
    Set SEL_MOBACTMST_QR = rdb.CreateQuery("", str)
    Set SEL_MOBACTMST_RS = SEL_MOBACTMST_QR.OpenResultset(1, 3)
    While Not SEL_MOBACTMST_RS.EOF
        ListBacteriaG409.AddItem Trim(SEL_MOBACTMST_RS("BACTERIA"))
        SEL_MOBACTMST_RS.MoveNext
    Wend
    
    '*****
    str = "select * FROM DISLIB.MOYEASTMST WHERE SER_CODE='G409' ORDER BY YEAST"
    Set SEL_MOYEASTMST_QR = rdb.CreateQuery("", str)
    Set SEL_MOYEASTMST_RS = SEL_MOYEASTMST_QR.OpenResultset(1, 3)
    While Not SEL_MOYEASTMST_RS.EOF
        ListYeastG409.AddItem Trim(SEL_MOYEASTMST_RS("YEAST"))
        SEL_MOYEASTMST_RS.MoveNext
    Wend
    SSTab_MO.Tab = 1
ElseIf SEL_DMUL_RS("PRO_CODE") = "G406" Then

    SSTab_MO.TabEnabled(2) = True
    SSTab_MO.TabVisible(2) = True
    cmb_406virDet.Enabled = True
    cmb_406BactDet.Enabled = True
    cmb_406ParaDet.Enabled = True
    
    str = "select * FROM DISLIB.MOVIRUSMST WHERE SER_CODE='G406' ORDER BY VIRUS"
    Set SEL_MOVIRMST_QR = rdb.CreateQuery("", str)
    Set SEL_MOVIRMST_RS = SEL_MOVIRMST_QR.OpenResultset(1, 3)
    While Not SEL_MOVIRMST_RS.EOF
        ListVirusG406.AddItem Trim(SEL_MOVIRMST_RS("VIRUS"))
        SEL_MOVIRMST_RS.MoveNext
    Wend
    
    '*****
    
    str = "select * FROM DISLIB.MOBACTEMST WHERE SER_CODE='G406' ORDER BY BACTERIA"
    Set SEL_MOBACTMST_QR = rdb.CreateQuery("", str)
    Set SEL_MOBACTMST_RS = SEL_MOBACTMST_QR.OpenResultset(1, 3)
    While Not SEL_MOBACTMST_RS.EOF
        ListBacteriaG406.AddItem Trim(SEL_MOBACTMST_RS("BACTERIA"))
        SEL_MOBACTMST_RS.MoveNext
    Wend
    
    '*****
    str = "select * FROM DISLIB.MOPARASMST WHERE SER_CODE='G406' ORDER BY PARASITES"
    Set SEL_MOPARAMST_QR = rdb.CreateQuery("", str)
    Set SEL_MOPARAMST_RS = SEL_MOPARAMST_QR.OpenResultset(1, 3)
    While Not SEL_MOPARAMST_RS.EOF
        ListParaG406.AddItem Trim(SEL_MOPARAMST_RS("PARASITES"))
        SEL_MOPARAMST_RS.MoveNext
    Wend
    SSTab_MO.Tab = 2
ElseIf SEL_DMUL_RS("PRO_CODE") = "G410" Then

    SSTab_MO.TabEnabled(3) = True
    SSTab_MO.TabVisible(3) = True
    
    cmb_410virDet.Enabled = True
    cmb_410BactDet.Enabled = True
    cmb_410AntiDet.Enabled = True
    cmb_410AtypicalDet.Enabled = True
    
    str = "select * FROM DISLIB.MOVIRUSMST WHERE SER_CODE='G410' ORDER BY VIRUS"
    Set SEL_MOVIRMST_QR = rdb.CreateQuery("", str)
    Set SEL_MOVIRMST_RS = SEL_MOVIRMST_QR.OpenResultset(1, 3)
    While Not SEL_MOVIRMST_RS.EOF
        ListVirusG410.AddItem Trim(SEL_MOVIRMST_RS("VIRUS"))
        SEL_MOVIRMST_RS.MoveNext
    Wend
    
    '*****
    
    str = "select * FROM DISLIB.MOBACTEMST WHERE SER_CODE='G410' ORDER BY BACTERIA"
    Set SEL_MOBACTMST_QR = rdb.CreateQuery("", str)
    Set SEL_MOBACTMST_RS = SEL_MOBACTMST_QR.OpenResultset(1, 3)
    While Not SEL_MOBACTMST_RS.EOF
        ListBacteriaG410.AddItem Trim(SEL_MOBACTMST_RS("BACTERIA"))
        SEL_MOBACTMST_RS.MoveNext
    Wend
    
    '*****
    str = "select * FROM DISLIB.MORSGENMST WHERE SER_CODE='G410' ORDER BY RESGENES"
    Set SEL_MOANTIMST_QR = rdb.CreateQuery("", str)
    Set SEL_MOANTIMST_RS = SEL_MOANTIMST_QR.OpenResultset(1, 3)
    While Not SEL_MOANTIMST_RS.EOF
        ListAntiG410.AddItem Trim(SEL_MOANTIMST_RS("RESGENES"))
        SEL_MOANTIMST_RS.MoveNext
    Wend
  '*****
    str = "select * FROM DISLIB.MOATYPBMST WHERE SER_CODE='G410' ORDER BY BACTERIA"
    Set SEL_MOATYPMST_QR = rdb.CreateQuery("", str)
    Set SEL_MOATYPMST_RS = SEL_MOATYPMST_QR.OpenResultset(1, 3)
    While Not SEL_MOATYPMST_RS.EOF
        ListAtypicalG410.AddItem Trim(SEL_MOATYPMST_RS("BACTERIA"))
        SEL_MOATYPMST_RS.MoveNext
    Wend
    SSTab_MO.Tab = 3
ElseIf SEL_DMUL_RS("PRO_CODE") = "G407" Then

    SSTab_MO.TabEnabled(4) = True
    SSTab_MO.TabVisible(4) = True
    
    cmb_407GramPos.Enabled = True
    cmb_407GramNeg.Enabled = True
    cmb_407Anti.Enabled = True
    cmb_407Yeast.Enabled = True
    
    cmd_Go_Click
    
    str = "select * FROM DISLIB.MOGRAMPMST WHERE SER_CODE='G407' AND ACTIVEYN=? ORDER BY BACTERIA"
    Set SEL_MOBACTMST_QR = rdb.CreateQuery("", str)
    
    SEL_MOBACTMST_QR(0) = Trim(cmb_ACTIVEYN.Text)
    Set SEL_MOBACTMST_RS = SEL_MOBACTMST_QR.OpenResultset(1, 3)
    While Not SEL_MOBACTMST_RS.EOF
        ListG407Gpos.AddItem Trim(SEL_MOBACTMST_RS("BACTERIA"))
        SEL_MOBACTMST_RS.MoveNext
    Wend
    
    '*****
    
    str = "select * FROM DISLIB.MOGRAMNMST WHERE SER_CODE='G407' AND ACTIVEYN=? ORDER BY BACTERIA"
    Set SEL_MOBACTMST_QR = rdb.CreateQuery("", str)
    
    SEL_MOBACTMST_QR(0) = Trim(cmb_ACTIVEYN.Text)
    Set SEL_MOBACTMST_RS = SEL_MOBACTMST_QR.OpenResultset(1, 3)
    While Not SEL_MOBACTMST_RS.EOF
        ListG407Gneg.AddItem Trim(SEL_MOBACTMST_RS("BACTERIA"))
        SEL_MOBACTMST_RS.MoveNext
    Wend
    
    '*****
    str = "select * FROM DISLIB.MORSGENMST WHERE SER_CODE='G407' AND ACTIVEYN=? ORDER BY RESGENES"
    Set SEL_MOANTIMST_QR = rdb.CreateQuery("", str)
    
    SEL_MOANTIMST_QR(0) = Trim(cmb_ACTIVEYN.Text)
    Set SEL_MOANTIMST_RS = SEL_MOANTIMST_QR.OpenResultset(1, 3)
    While Not SEL_MOANTIMST_RS.EOF
        ListG407Anti.AddItem Trim(SEL_MOANTIMST_RS("RESGENES"))
        SEL_MOANTIMST_RS.MoveNext
    Wend
  '*****
    str = "select * FROM DISLIB.MOYEASTMST WHERE SER_CODE='G407' AND ACTIVEYN=? ORDER BY YEAST"
    Set SEL_MOYEASTMST_QR = rdb.CreateQuery("", str)
    
    SEL_MOYEASTMST_QR(0) = Trim(cmb_ACTIVEYN.Text)
    Set SEL_MOYEASTMST_RS = SEL_MOYEASTMST_QR.OpenResultset(1, 3)
    While Not SEL_MOYEASTMST_RS.EOF
        ListG407Yeast.AddItem Trim(SEL_MOYEASTMST_RS("YEAST"))
        SEL_MOYEASTMST_RS.MoveNext
    Wend
    SSTab_MO.Tab = 4
    
End If
'*****

End If

If viewflag = True Then
    
    SSTab_MO.TabEnabled(0) = False
    SSTab_MO.TabEnabled(1) = False
    SSTab_MO.TabEnabled(2) = False
    SSTab_MO.TabEnabled(3) = False
    SSTab_MO.TabEnabled(4) = False
    
'    SSTab_MO.TabVisible(0) = False
'    SSTab_MO.TabVisible(1) = False
'    SSTab_MO.TabVisible(2) = False
'    SSTab_MO.TabVisible(3) = False
'    SSTab_MO.TabVisible(4) = False
    
    
    SEL_DMUL_QR(0) = Trim(cmb_reqno.Text)
    Set SEL_DMUL_RS = SEL_DMUL_QR.OpenResultset(1, 3)
    
    If SEL_DMUL_RS("PRO_CODE") = "G408" Then
        SSTab_MO.TabEnabled(0) = True
        SSTab_MO.TabVisible(0) = True
        cmb_408virDet.Enabled = True
        cmb_408BactDet.Enabled = True

        str = "select * FROM DISLIB.MOVIRUSMST WHERE SER_CODE='G408' ORDER BY VIRUS"
        Set SEL_MOVIRMST_QR = rdb.CreateQuery("", str)
        Set SEL_MOVIRMST_RS = SEL_MOVIRMST_QR.OpenResultset(1, 3)
        While Not SEL_MOVIRMST_RS.EOF
            ListVirusG408.AddItem Trim(SEL_MOVIRMST_RS("VIRUS"))
            SEL_MOVIRMST_RS.MoveNext
        Wend
        
        '*****
        
        str = "select * FROM DISLIB.MOBACTEMST WHERE SER_CODE='G408' ORDER BY BACTERIA"
        Set SEL_MOBACTMST_QR = rdb.CreateQuery("", str)
        Set SEL_MOBACTMST_RS = SEL_MOBACTMST_QR.OpenResultset(1, 3)
        While Not SEL_MOBACTMST_RS.EOF
            ListBacteriaG408.AddItem Trim(SEL_MOBACTMST_RS("BACTERIA"))
            SEL_MOBACTMST_RS.MoveNext
        Wend
        
        str = "SELECT * FROM DISLIB.G408TRN WHERE reqno = ? "
        Set SEL_408TRN_QR = rdb.CreateQuery("", str)

        SEL_408TRN_QR(0) = Trim(cmb_reqno.Text)
        Set SEL_408TRN_RS = SEL_408TRN_QR.OpenResultset(1, 3)
        If SEL_408TRN_RS.EOF = False Then
               
        Txt_VirusG408.Text = Trim(SEL_408TRN_RS("VIRUS"))
        Txt_BacteriaG408.Text = Trim(SEL_408TRN_RS("BACTERIA"))
        txt_MOremarks.Text = Trim(SEL_408TRN_RS("REMARKS"))
        End If
            SSTab_MO.Tab = 0
    ElseIf SEL_DMUL_RS("PRO_CODE") = "G409" Then
        SSTab_MO.TabEnabled(1) = True
        SSTab_MO.TabVisible(1) = True
        cmb_409virDet.Enabled = True
        cmb_409BactDet.Enabled = True
        cmb_409YeastDet.Enabled = True
        
        str = "select * FROM DISLIB.MOVIRUSMST WHERE SER_CODE='G409' ORDER BY VIRUS"
        Set SEL_MOVIRMST_QR = rdb.CreateQuery("", str)
        Set SEL_MOVIRMST_RS = SEL_MOVIRMST_QR.OpenResultset(1, 3)
        While Not SEL_MOVIRMST_RS.EOF
            ListVirusG409.AddItem Trim(SEL_MOVIRMST_RS("VIRUS"))
            SEL_MOVIRMST_RS.MoveNext
        Wend
        
        '*****
        
        str = "select * FROM DISLIB.MOBACTEMST WHERE SER_CODE='G409' ORDER BY BACTERIA"
        Set SEL_MOBACTMST_QR = rdb.CreateQuery("", str)
        Set SEL_MOBACTMST_RS = SEL_MOBACTMST_QR.OpenResultset(1, 3)
        While Not SEL_MOBACTMST_RS.EOF
            ListBacteriaG409.AddItem Trim(SEL_MOBACTMST_RS("BACTERIA"))
            SEL_MOBACTMST_RS.MoveNext
        Wend
        
        '*****
        '*****
        str = "select * FROM DISLIB.MOYEASTMST WHERE SER_CODE='G409' ORDER BY YEAST"
        Set SEL_MOYEASTMST_QR = rdb.CreateQuery("", str)
        Set SEL_MOYEASTMST_RS = SEL_MOYEASTMST_QR.OpenResultset(1, 3)
        While Not SEL_MOYEASTMST_RS.EOF
            ListYeastG409.AddItem Trim(SEL_MOYEASTMST_RS("YEAST"))
            SEL_MOYEASTMST_RS.MoveNext
        Wend
        
        str = "SELECT * FROM DISLIB.G409TRN WHERE reqno = ? "
        Set SEL_409TRN_QR = rdb.CreateQuery("", str)

        SEL_409TRN_QR(0) = Trim(cmb_reqno.Text)
        Set SEL_409TRN_RS = SEL_409TRN_QR.OpenResultset(1, 3)
        If SEL_409TRN_RS.EOF = False Then
        Txt_VirusG409.Text = Trim(SEL_409TRN_RS("VIRUS"))
        Txt_BacteriaG409.Text = Trim(SEL_409TRN_RS("BACTERIA"))
        Txt_YeastG409.Text = Trim(SEL_409TRN_RS("YEAST"))
        txt_MOremarks.Text = Trim(SEL_409TRN_RS("REMARKS"))
        End If
            SSTab_MO.Tab = 1
    ElseIf SEL_DMUL_RS("PRO_CODE") = "G406" Then
        SSTab_MO.TabEnabled(2) = True
        SSTab_MO.TabVisible(2) = True
        cmb_406virDet.Enabled = True
        cmb_406BactDet.Enabled = True
        cmb_406ParaDet.Enabled = True

        str = "select * FROM DISLIB.MOVIRUSMST WHERE SER_CODE='G406' ORDER BY VIRUS"
        Set SEL_MOVIRMST_QR = rdb.CreateQuery("", str)
        Set SEL_MOVIRMST_RS = SEL_MOVIRMST_QR.OpenResultset(1, 3)
        While Not SEL_MOVIRMST_RS.EOF
            ListVirusG406.AddItem Trim(SEL_MOVIRMST_RS("VIRUS"))
            SEL_MOVIRMST_RS.MoveNext
        Wend
        
        '*****
        
        str = "select * FROM DISLIB.MOBACTEMST WHERE SER_CODE='G406' ORDER BY BACTERIA"
        Set SEL_MOBACTMST_QR = rdb.CreateQuery("", str)
        Set SEL_MOBACTMST_RS = SEL_MOBACTMST_QR.OpenResultset(1, 3)
        While Not SEL_MOBACTMST_RS.EOF
            ListBacteriaG406.AddItem Trim(SEL_MOBACTMST_RS("BACTERIA"))
            SEL_MOBACTMST_RS.MoveNext
        Wend
        
        '*****
        '*****
        str = "select * FROM DISLIB.MOPARASMST WHERE SER_CODE='G406' ORDER BY parasites"
        Set SEL_MOPARAMST_QR = rdb.CreateQuery("", str)
        Set SEL_MOPARAMST_RS = SEL_MOPARAMST_QR.OpenResultset(1, 3)
        While Not SEL_MOPARAMST_RS.EOF
            ListParaG406.AddItem Trim(SEL_MOPARAMST_RS("PARASITES"))
            SEL_MOPARAMST_RS.MoveNext
        Wend
        
        str = "SELECT * FROM DISLIB.G406TRN WHERE reqno = ? "
        Set SEL_406TRN_QR = rdb.CreateQuery("", str)
      
        SEL_406TRN_QR(0) = Trim(cmb_reqno.Text)
        Set SEL_406TRN_RS = SEL_406TRN_QR.OpenResultset(1, 3)
        If SEL_406TRN_RS.EOF = False Then
        Txt_VirusG406.Text = Trim(SEL_406TRN_RS("VIRUS"))
        Txt_BacteriaG406.Text = Trim(SEL_406TRN_RS("BACTERIA"))
        Txt_ParaG406.Text = Trim(SEL_406TRN_RS("PARA"))
        txt_MOremarks.Text = Trim(SEL_406TRN_RS("REMARKS"))
        End If
            SSTab_MO.Tab = 2
    ElseIf SEL_DMUL_RS("PRO_CODE") = "G410" Then
        SSTab_MO.TabEnabled(3) = True
        SSTab_MO.TabVisible(3) = True
        
        cmb_410virDet.Enabled = True
        cmb_410BactDet.Enabled = True
        cmb_410AntiDet.Enabled = True
        cmb_410AtypicalDet.Enabled = True
        
        str = "select * FROM DISLIB.MOVIRUSMST WHERE SER_CODE='G410' ORDER BY VIRUS"
        Set SEL_MOVIRMST_QR = rdb.CreateQuery("", str)
        Set SEL_MOVIRMST_RS = SEL_MOVIRMST_QR.OpenResultset(1, 3)
        While Not SEL_MOVIRMST_RS.EOF
            ListVirusG410.AddItem Trim(SEL_MOVIRMST_RS("VIRUS"))
            SEL_MOVIRMST_RS.MoveNext
        Wend
        
        '*****
        
        str = "select * FROM DISLIB.MOBACTEMST WHERE SER_CODE='G410' ORDER BY BACTERIA"
        Set SEL_MOBACTMST_QR = rdb.CreateQuery("", str)
        Set SEL_MOBACTMST_RS = SEL_MOBACTMST_QR.OpenResultset(1, 3)
        While Not SEL_MOBACTMST_RS.EOF
            ListBacteriaG410.AddItem Trim(SEL_MOBACTMST_RS("BACTERIA"))
            SEL_MOBACTMST_RS.MoveNext
        Wend
        
        '*****
        str = "select * FROM DISLIB.MORSGENMST WHERE SER_CODE='G410' ORDER BY RESGENES"
        Set SEL_MOANTIMST_QR = rdb.CreateQuery("", str)
        Set SEL_MOANTIMST_RS = SEL_MOANTIMST_QR.OpenResultset(1, 3)
        While Not SEL_MOANTIMST_RS.EOF
            ListAntiG410.AddItem Trim(SEL_MOANTIMST_RS("RESGENES"))
            SEL_MOANTIMST_RS.MoveNext
        Wend
        '*****
        str = "select * FROM DISLIB.MOATYPBMST WHERE SER_CODE='G410' ORDER BY BACTERIA"
        Set SEL_MOATYPMST_QR = rdb.CreateQuery("", str)
        Set SEL_MOATYPMST_RS = SEL_MOATYPMST_QR.OpenResultset(1, 3)
        While Not SEL_MOATYPMST_RS.EOF
            ListAtypicalG410.AddItem Trim(SEL_MOATYPMST_RS("BACTERIA"))
            SEL_MOATYPMST_RS.MoveNext
        Wend
        
        str = "SELECT * FROM DISLIB.G410TRN WHERE reqno = ? "
        Set SEL_410TRN_QR = rdb.CreateQuery("", str)

        SEL_410TRN_QR(0) = Trim(cmb_reqno.Text)
        Set SEL_410TRN_RS = SEL_410TRN_QR.OpenResultset(1, 3)
        If SEL_410TRN_RS.EOF = False Then
        
        Txt_VirusG410.Text = Trim(SEL_410TRN_RS("VIRUS"))
        Txt_BacteriaG410.Text = Trim(SEL_410TRN_RS("BACTERIA"))
        Txt_AntiG410.Text = Trim(SEL_410TRN_RS("ANTI"))
        Txt_AtypicalG410.Text = Trim(SEL_410TRN_RS("ATYPICAL"))
        txt_MOremarks.Text = Trim(SEL_410TRN_RS("REMARKS"))
        End If
            SSTab_MO.Tab = 3
    ElseIf SEL_DMUL_RS("PRO_CODE") = "G407" Then
        SSTab_MO.TabEnabled(4) = True
        SSTab_MO.TabVisible(4) = True
        cmb_407GramPos.Enabled = True
        cmb_407GramNeg.Enabled = True
        cmb_407Anti.Enabled = True
        cmb_407Yeast.Enabled = True


        cmd_Go_Click
        
        str = "select * FROM DISLIB.MOGRAMPMST WHERE SER_CODE='G407' AND ACTIVEYN=? ORDER BY BACTERIA"
        Set SEL_MOBACTMST_QR = rdb.CreateQuery("", str)
        
        SEL_MOBACTMST_QR(0) = Trim(cmb_ACTIVEYN.Text)
        Set SEL_MOBACTMST_RS = SEL_MOBACTMST_QR.OpenResultset(1, 3)
        While Not SEL_MOBACTMST_RS.EOF
            ListG407Gpos.AddItem Trim(SEL_MOBACTMST_RS("BACTERIA"))
            SEL_MOBACTMST_RS.MoveNext
        Wend
        
        '*****
        
        str = "select * FROM DISLIB.MOGRAMNMST WHERE SER_CODE='G407' AND ACTIVEYN=? ORDER BY BACTERIA"
        Set SEL_MOBACTMST_QR = rdb.CreateQuery("", str)
        
        SEL_MOBACTMST_QR(0) = Trim(cmb_ACTIVEYN.Text)
        Set SEL_MOBACTMST_RS = SEL_MOBACTMST_QR.OpenResultset(1, 3)
        While Not SEL_MOBACTMST_RS.EOF
            ListG407Gneg.AddItem Trim(SEL_MOBACTMST_RS("BACTERIA"))
            SEL_MOBACTMST_RS.MoveNext
        Wend
        
        '*****
        str = "select * FROM DISLIB.MORSGENMST WHERE SER_CODE='G407' AND ACTIVEYN=? ORDER BY RESGENES"
        Set SEL_MOANTIMST_QR = rdb.CreateQuery("", str)
        
        SEL_MOANTIMST_QR(0) = Trim(cmb_ACTIVEYN.Text)
        Set SEL_MOANTIMST_RS = SEL_MOANTIMST_QR.OpenResultset(1, 3)
        While Not SEL_MOANTIMST_RS.EOF
            ListG407Anti.AddItem Trim(SEL_MOANTIMST_RS("RESGENES"))
            SEL_MOANTIMST_RS.MoveNext
        Wend
        '*****
        str = "select * FROM DISLIB.MOYEASTMST WHERE SER_CODE='G407' AND ACTIVEYN=? ORDER BY YEAST"
        Set SEL_MOYEASTMST_QR = rdb.CreateQuery("", str)
        
        SEL_MOYEASTMST_QR(0) = Trim(cmb_ACTIVEYN.Text)
        Set SEL_MOYEASTMST_RS = SEL_MOYEASTMST_QR.OpenResultset(1, 3)
        While Not SEL_MOYEASTMST_RS.EOF
            ListG407Yeast.AddItem Trim(SEL_MOYEASTMST_RS("YEAST"))
            SEL_MOYEASTMST_RS.MoveNext
        Wend
        
        str = "SELECT * FROM DISLIB.G407TRN WHERE reqno = ? "
        Set SEL_407TRN_QR = rdb.CreateQuery("", str)

        SEL_407TRN_QR(0) = Trim(cmb_reqno.Text)
        Set SEL_407TRN_RS = SEL_407TRN_QR.OpenResultset(1, 3)
        If SEL_407TRN_RS.EOF Then
        
        Txt_407Gpos.Text = Trim(SEL_407TRN_RS("GPOS"))
        Txt_407Gneg.Text = Trim(SEL_407TRN_RS("GNEG"))
        Txt_407Anti.Text = Trim(SEL_407TRN_RS("ANTI"))
        Txt_407Yeast.Text = Trim(SEL_407TRN_RS("YEAST"))
        txt_MOremarks.Text = Trim(SEL_407TRN_RS("REMARKS"))
        cmb_ACTIVEYN.Text = Trim(SEL_407TRN_RS("ACTIVEYN"))
        End If
            SSTab_MO.Tab = 4
    End If
    
    
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
    Set REFDR_QR = rdb.CreateQuery("", str)
    
    REFDR_QR(0) = Trim(cmb_reqno.Text)
    Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
    
    If REFDR_RS.RowCount > 0 Then
        
        If Trim(REFDR_RS("COMMITFLAG")) = "Y" And Trim(REFDR_RS("REQNSTAT")) = "F" Then
            cmb_finalby.Text = Trim(REFDR_RS("FREPORTBY"))
            MsgBox "Final Report is Generated. Hence no more modifications."
            cmd_modify.Enabled = False
            cmd_save.Enabled = False
            
       
            
            cmb_408virDet.Enabled = False
            cmb_408BactDet.Enabled = False
            
            cmb_409virDet.Enabled = False
            cmb_409BactDet.Enabled = False
            cmb_409YeastDet.Enabled = False
            
            cmb_406virDet.Enabled = False
            cmb_406BactDet.Enabled = False
            cmb_406ParaDet.Enabled = False
            
            cmb_410virDet.Enabled = False
            cmb_410BactDet.Enabled = False
            cmb_410AntiDet.Enabled = False
            cmb_410AtypicalDet.Enabled = False
            
            cmb_407GramPos.Enabled = False
            cmb_407GramNeg.Enabled = False
            cmb_407Anti.Enabled = False
            cmb_407Yeast.Enabled = False
            '*******************************
            
        ElseIf (Trim(REFDR_RS("COMMITFLAG")) = "" Or IsNull(REFDR_RS("COMMITFLAG")) = True) And Trim(REFDR_RS("REQNSTAT")) = "T" Then
            cmd_modify.Enabled = True
        End If
    End If
    
End If
End Sub

Public Sub assign()

On Error GoTo MSG:

rdb.BeginTrans

If flagsave = 1 Then

    If SSTab_MO.TabEnabled(0) = True Then
        str = "insert into DISLIB.G408TRN(REQNO,VIRUS,BACTERIA,REMARKS)values(?,?,?,?)"
        Set INS_408TRN_QR = rdb.CreateQuery("", str)
    
        INS_408TRN_QR(0) = Trim(cmb_reqno.Text)
        INS_408TRN_QR(1) = Trim(Txt_VirusG408.Text)
        INS_408TRN_QR(2) = Trim(Txt_BacteriaG408.Text)
        INS_408TRN_QR(3) = Trim(txt_MOremarks.Text)
        INS_408TRN_QR.Execute
    
        '-- Update rislib.rrequisitn
        'update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=? WHERE REQNNO=?"
        upd_Trrequisitn_qr(0) = cmb_enteryby.Text
        'upd_Trrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        upd_Trrequisitn_qr(1) = cmb_finalby.Text
        upd_Trrequisitn_qr(2) = "T"
        upd_Trrequisitn_qr(3) = Trim(cmb_reqno.Text)
        upd_Trrequisitn_qr.Execute
        
        '-- Update rislib.rTIMELOG
        'PDRTIME_QR(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        'UPDRTIME_QR(1) = time1
        UPDRTIME_QR(0) = Trim(cmb_reqno.Text)
        UPDRTIME_QR.Execute
    ElseIf SSTab_MO.TabEnabled(1) = True Then
        str = "insert into DISLIB.G409TRN(REQNO,VIRUS,BACTERIA,YEAST,REMARKS)values(?,?,?,?,?)"
        Set INS_409TRN_QR = rdb.CreateQuery("", str)
    
        INS_409TRN_QR(0) = Trim(cmb_reqno.Text)
        INS_409TRN_QR(1) = Trim(Txt_VirusG409.Text)
        INS_409TRN_QR(2) = Trim(Txt_BacteriaG409.Text)
        INS_409TRN_QR(3) = Trim(Txt_YeastG409.Text)
        INS_409TRN_QR(4) = Trim(txt_MOremarks.Text)
        INS_409TRN_QR.Execute
    
        '-- Update rislib.rrequisitn
        'update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=? WHERE REQNNO=?"
        upd_Trrequisitn_qr(0) = cmb_enteryby.Text
        'upd_Trrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        upd_Trrequisitn_qr(1) = cmb_finalby.Text
        upd_Trrequisitn_qr(2) = "T"
        upd_Trrequisitn_qr(3) = Trim(cmb_reqno.Text)
        upd_Trrequisitn_qr.Execute
        
        '-- Update rislib.rTIMELOG
        'PDRTIME_QR(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        'UPDRTIME_QR(1) = time1
        UPDRTIME_QR(0) = Trim(cmb_reqno.Text)
        UPDRTIME_QR.Execute
    
    ElseIf SSTab_MO.TabEnabled(2) = True Then
        str = "insert into DISLIB.G406TRN(REQNO,VIRUS,BACTERIA,PARA,REMARKS)values(?,?,?,?,?)"
        Set INS_406TRN_QR = rdb.CreateQuery("", str)
    
        INS_406TRN_QR(0) = Trim(cmb_reqno.Text)
        INS_406TRN_QR(1) = Trim(Txt_VirusG406.Text)
        INS_406TRN_QR(2) = Trim(Txt_BacteriaG406.Text)
        INS_406TRN_QR(3) = Trim(Txt_ParaG406.Text)
        INS_406TRN_QR(4) = Trim(txt_MOremarks.Text)
        INS_406TRN_QR.Execute
    
        
        upd_Trrequisitn_qr(0) = cmb_enteryby.Text
        upd_Trrequisitn_qr(1) = cmb_finalby.Text
        upd_Trrequisitn_qr(2) = "T"
        upd_Trrequisitn_qr(3) = Trim(cmb_reqno.Text)
        upd_Trrequisitn_qr.Execute
                
        UPDRTIME_QR(0) = Trim(cmb_reqno.Text)
        UPDRTIME_QR.Execute
    
    ElseIf SSTab_MO.TabEnabled(3) = True Then
        str = "insert into DISLIB.G410TRN(REQNO,VIRUS,BACTERIA,ANTI,ATYPICAL,REMARKS)values(?,?,?,?,?,?)"
        Set INS_410TRN_QR = rdb.CreateQuery("", str)
    
        INS_410TRN_QR(0) = Trim(cmb_reqno.Text)
        INS_410TRN_QR(1) = Trim(Txt_VirusG410.Text)
        INS_410TRN_QR(2) = Trim(Txt_BacteriaG410.Text)
        INS_410TRN_QR(3) = Trim(Txt_AntiG410.Text)
        INS_410TRN_QR(4) = Trim(Txt_AtypicalG410.Text)
        INS_410TRN_QR(5) = Trim(txt_MOremarks.Text)
        INS_410TRN_QR.Execute
    
        
        upd_Trrequisitn_qr(0) = cmb_enteryby.Text
        upd_Trrequisitn_qr(1) = cmb_finalby.Text
        upd_Trrequisitn_qr(2) = "T"
        upd_Trrequisitn_qr(3) = Trim(cmb_reqno.Text)
        upd_Trrequisitn_qr.Execute
                
        UPDRTIME_QR(0) = Trim(cmb_reqno.Text)
        UPDRTIME_QR.Execute
        
    ElseIf SSTab_MO.TabEnabled(4) = True Then
        str = "insert into DISLIB.G407TRN(REQNO,GPOS,GNEG,ANTI,YEAST,REMARKS,ACTIVEYN)values(?,?,?,?,?,?,?)"
        Set INS_407TRN_QR = rdb.CreateQuery("", str)
    
        INS_407TRN_QR(0) = Trim(cmb_reqno.Text)
        INS_407TRN_QR(1) = Trim(Txt_407Gpos.Text)
        INS_407TRN_QR(2) = Trim(Txt_407Gneg.Text)
        INS_407TRN_QR(3) = Trim(Txt_407Anti.Text)
        INS_407TRN_QR(4) = Trim(Txt_407Yeast.Text)
        INS_407TRN_QR(5) = Trim(txt_MOremarks.Text)
        INS_407TRN_QR(6) = Trim(cmb_ACTIVEYN.Text)
        INS_407TRN_QR.Execute
    
        
        upd_Trrequisitn_qr(0) = cmb_enteryby.Text
        upd_Trrequisitn_qr(1) = cmb_finalby.Text
        upd_Trrequisitn_qr(2) = "T"
        upd_Trrequisitn_qr(3) = Trim(cmb_reqno.Text)
        upd_Trrequisitn_qr.Execute
                
        UPDRTIME_QR(0) = Trim(cmb_reqno.Text)
        UPDRTIME_QR.Execute
    End If
    flagsave = 0
    
    
End If

If flagsave = 2 Then
    If SSTab_MO.TabEnabled(0) = True Then
        str = "update dislib.G408TRN set VIRUS=? ,BACTERIA=?,REMARKS=? where reqno=? "
        Set UPD_408TRN_QR = rdb.CreateQuery("", str)
        
        UPD_408TRN_QR(0) = Trim(Txt_VirusG408.Text)
        UPD_408TRN_QR(1) = Trim(Txt_BacteriaG408.Text)
        UPD_408TRN_QR(2) = Trim(txt_MOremarks.Text)
        UPD_408TRN_QR(3) = Trim(cmb_reqno.Text)
        UPD_408TRN_QR.Execute
        
        '-- Update rislib.rrequisitn
        'update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=? WHERE REQNNO=?"
        upd_Trrequisitn_qr(0) = cmb_enteryby.Text
        upd_Trrequisitn_qr(1) = cmb_finalby.Text
        upd_Trrequisitn_qr(2) = "T"
        upd_Trrequisitn_qr(3) = Trim(cmb_reqno.Text)
        upd_Trrequisitn_qr.Execute
        
        '-- Update rislib.rTIMELOG
        UPDRTIME_QR(0) = Trim(cmb_reqno.Text)
        UPDRTIME_QR.Execute
    ElseIf SSTab_MO.TabEnabled(1) = True Then
        str = "update dislib.G409TRN set VIRUS=? ,BACTERIA=?,YEAST=?,REMARKS=? where reqno=? "
        Set UPD_409TRN_QR = rdb.CreateQuery("", str)
        
        UPD_409TRN_QR(0) = Trim(Txt_VirusG409.Text)
        UPD_409TRN_QR(1) = Trim(Txt_BacteriaG409.Text)
        UPD_409TRN_QR(2) = Trim(Txt_YeastG409.Text)
        UPD_409TRN_QR(3) = Trim(txt_MOremarks.Text)
        UPD_409TRN_QR(4) = Trim(cmb_reqno.Text)
        UPD_409TRN_QR.Execute
        
        '-- Update rislib.rrequisitn
        'update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=? WHERE REQNNO=?"
        upd_Trrequisitn_qr(0) = cmb_enteryby.Text
        upd_Trrequisitn_qr(1) = cmb_finalby.Text
        upd_Trrequisitn_qr(2) = "T"
        upd_Trrequisitn_qr(3) = Trim(cmb_reqno.Text)
        upd_Trrequisitn_qr.Execute
        
        '-- Update rislib.rTIMELOG
        UPDRTIME_QR(0) = Trim(cmb_reqno.Text)
        UPDRTIME_QR.Execute
    ElseIf SSTab_MO.TabEnabled(2) = True Then
        str = "update dislib.G406TRN set VIRUS=? ,BACTERIA=?,PARA=?,REMARKS=? where reqno=? "
        Set UPD_406TRN_QR = rdb.CreateQuery("", str)
        
        UPD_406TRN_QR(0) = Trim(Txt_VirusG406.Text)
        UPD_406TRN_QR(1) = Trim(Txt_BacteriaG406.Text)
        UPD_406TRN_QR(2) = Trim(Txt_ParaG406.Text)
        UPD_406TRN_QR(3) = Trim(txt_MOremarks.Text)
        UPD_406TRN_QR(4) = Trim(cmb_reqno.Text)
        UPD_406TRN_QR.Execute
       
        upd_Trrequisitn_qr(0) = cmb_enteryby.Text
        upd_Trrequisitn_qr(1) = cmb_finalby.Text
        upd_Trrequisitn_qr(2) = "T"
        upd_Trrequisitn_qr(3) = Trim(cmb_reqno.Text)
        upd_Trrequisitn_qr.Execute
                
        UPDRTIME_QR(0) = Trim(cmb_reqno.Text)
        UPDRTIME_QR.Execute
    ElseIf SSTab_MO.TabEnabled(3) = True Then
        str = "update dislib.G410TRN set VIRUS=? ,BACTERIA=?,ANTI=?, ATYPICAL=?,REMARKS=? where reqno=? "
        Set UPD_410TRN_QR = rdb.CreateQuery("", str)
        
        UPD_410TRN_QR(0) = Trim(Txt_VirusG410.Text)
        UPD_410TRN_QR(1) = Trim(Txt_BacteriaG410.Text)
        UPD_410TRN_QR(2) = Trim(Txt_AntiG410.Text)
        UPD_410TRN_QR(3) = Trim(Txt_AtypicalG410.Text)
        UPD_410TRN_QR(4) = Trim(txt_MOremarks.Text)
        UPD_410TRN_QR(5) = Trim(cmb_reqno.Text)
        UPD_410TRN_QR.Execute
       
        upd_Trrequisitn_qr(0) = cmb_enteryby.Text
        upd_Trrequisitn_qr(1) = cmb_finalby.Text
        upd_Trrequisitn_qr(2) = "T"
        upd_Trrequisitn_qr(3) = Trim(cmb_reqno.Text)
        upd_Trrequisitn_qr.Execute
                
        UPDRTIME_QR(0) = Trim(cmb_reqno.Text)
        UPDRTIME_QR.Execute
    ElseIf SSTab_MO.TabEnabled(4) = True Then
        str = "update dislib.G407TRN set GPOS=? ,GNEG=?,ANTI=?, YEAST=?,REMARKS=?,ACTIVEYN=? where reqno=? "
        Set UPD_407TRN_QR = rdb.CreateQuery("", str)
        
        UPD_407TRN_QR(0) = Trim(Txt_407Gpos.Text)
        UPD_407TRN_QR(1) = Trim(Txt_407Gneg.Text)
        UPD_407TRN_QR(2) = Trim(Txt_407Anti.Text)
        UPD_407TRN_QR(3) = Trim(Txt_407Yeast.Text)
        UPD_407TRN_QR(4) = Trim(txt_MOremarks.Text)
        UPD_407TRN_QR(5) = Trim(cmb_ACTIVEYN.Text)
        UPD_407TRN_QR(6) = Trim(cmb_reqno.Text)
        UPD_407TRN_QR.Execute
       
        upd_Trrequisitn_qr(0) = cmb_enteryby.Text
        upd_Trrequisitn_qr(1) = cmb_finalby.Text
        upd_Trrequisitn_qr(2) = "T"
        upd_Trrequisitn_qr(3) = Trim(cmb_reqno.Text)
        upd_Trrequisitn_qr.Execute
                
        UPDRTIME_QR(0) = Trim(cmb_reqno.Text)
        UPDRTIME_QR.Execute
    End If
    flagsave = 0
End If

rdb.CommitTrans
MsgBox "Record saved Successfully."
  
BTNENABL1 ("SAVE")

Exit Sub
MSG:
    rdb.RollbackTrans
    MsgBox Err.Description

End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
If Opt_patient.Value = True Then
'If Trim(MaskEdBox1.Text) = "" Or MaskEdBox1.Text = "__/_____" Or MaskEdBox1.Text = "_______/_______" Then
If Trim(MaskEdBox1.Text) = "" Or MaskEdBox1.Text = "__/_____" Or MaskEdBox1.Text = "_______/_______" Then

    If (KeyAscii >= 49 And KeyAscii <= 58) Then
        MaskEdBox1.Mask = ">##?####/######"
    Else
        MaskEdBox1.Mask = ">??/#####"
    End If
End If
End If
KeyAscii = KeyAscii

End Sub

Private Sub Opt_patient_Click()
If Opt_patient.Value = True Then
    lbl_caseno.Caption = "Case No."
    MaskEdBox1.Mask = "??>/#####"
    MaskEdBox1.Enabled = True
    lbl_dmg.Visible = True
    txt_DMG.Visible = True
End If
End Sub

Private Sub opt_staff_Click()
If opt_staff.Value = True Then
    lbl_caseno.Caption = "CC No."
    MaskEdBox1.Mask = "######"
    MaskEdBox1.Enabled = True
    lbl_dmg.Visible = False
    txt_DMG.Visible = False
End If
End Sub


Private Sub MaskEdBox1_LostFocus()
Dim CASENO As String
If Len(Trim(MaskEdBox1.Text)) = "" Then
    Exit Sub
End If
CASENO = MaskEdBox1.Text
If Opt_patient.Value = True Then
  '  MaskEdBox1.Text = UCase(padzero(5, "C", CASENO, 3))
   MaskEdBox1.Text = NEWNOpad(CASENO)
End If
    
If NEWFLAG = True Then
    
    If MaskEdBox1.Text <> "" Then
        cmb_reqno.clear
        str = "select * from rislib.RREQUISITN where caseno=? AND REQNSTAT='E' AND MODALITYCD='MO' ORDER BY REQNDT desc"
        Set req_qr = rdb.CreateQuery("", str)
    
        req_qr(0) = MaskEdBox1.FormattedText
        Set req_rs = req_qr.OpenResultset(1, 3)
        If Not req_rs.EOF Then
            While Not req_rs.EOF
                cmb_reqno.AddItem Trim(req_rs("reqNno"))
                req_rs.MoveNext
            Wend
        End If
    End If
 End If
 
If viewflag = True Then
    cmb_reqno.clear
    str = "select * from rislib.RREQUISITN where caseno=? AND REQNSTAT IN ('T','F') AND MODALITYCD='MO' ORDER BY REQNDT desc"
    Set sel_Vreq_qr = rdb.CreateQuery("", str)
    
    sel_Vreq_qr(0) = MaskEdBox1.FormattedText
    Set sel_Vreq_rs = sel_Vreq_qr.OpenResultset(1, 3)

    If Not sel_Vreq_rs.EOF Then
        While Not sel_Vreq_rs.EOF
            cmb_reqno.AddItem Trim(sel_Vreq_rs("REQNNO"))
            sel_Vreq_rs.MoveNext
        Wend
    End If
    
End If

End Sub

Public Function padzero(l As Integer, ID As String, num As String, strpt As Integer) As String
Dim S
Dim ST1
Dim st2
Dim ST
Dim z
Dim z1
Dim c As Integer
Dim L1 As String
Dim l2 As String
If "num" Like "R-######" Or "num" Like "M-######" Or "num" Like ">??/#####" Or "num" Like ">??A/??/##/######" Then
    MsgBox ("format correct")
Else
    
    L1 = Left(num, 2) '**case no
    l2 = Left(num, 7) '** req no
     
    S = Mid(num, strpt)
    ST1 = Len(S)
    If ST1 < l Then
        z = l - ST1
        ST = 0
        z1 = z
        While z1 > 1
            ST = ST & "0"
            z1 = z1 - 1
        Wend
        S = ST & S
    End If
    ST = S
    num = ST
End If
    
   
    Select Case ID
    Case "R"
    st2 = "R-" & num
    Case "M"
    Set ZAS400DT = rdb.OpenResultset("SELECT DISTINCT CURRENT DATE DT1 , CURRENT TIME TM1 FROM PATOOLSLIB.MA107")
    st2 = "M" & Right(CDate(ZAS400DT("dt1")), 1) & Trim(num)
    ZAS400DT.Close
    Set ZAS400DT = Nothing
    Case "C"
    st2 = L1 & "/" & num
    Case "B"
    st2 = "B-" & num
    Case "F"
    st2 = "F-" & num
    Case "I"
    st2 = "I-" & num
    Case Is = "Q"
    st2 = l2 & num
    End Select

    padzero = st2
    
 
End Function

Private Sub BTNENABL1(cap As String)

If cap = "NEW" Then
    cmd_save.Enabled = True
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    CMD_NEW.Enabled = False
    cmd_modify.Enabled = False
    cmd_print.Enabled = False
    cmd_commit.Enabled = False
    cmd_view.Enabled = False

ElseIf cap = "VIEW" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_modify.Enabled = False
    CMD_NEW.Enabled = False
    cmd_save.Enabled = False
    cmd_print.Enabled = True
    cmd_commit.Enabled = False
    cmd_view.Enabled = False

ElseIf cap = "SAVE" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_modify.Enabled = False
    CMD_NEW.Enabled = True
    cmd_save.Enabled = False
    cmd_print.Enabled = True
    cmd_commit.Enabled = True
    cmd_view.Enabled = True

ElseIf cap = "INI" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_modify.Enabled = False
    CMD_NEW.Enabled = True
    cmd_save.Enabled = False
    cmd_print.Enabled = False
    cmd_commit.Enabled = False

    cmd_view.Enabled = True

ElseIf cap = "COMMIT" Then
    cmd_save.Enabled = False
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    CMD_NEW.Enabled = False
    cmd_modify.Enabled = False
    cmd_print.Enabled = True
    cmd_commit.Enabled = False
    cmd_view.Enabled = False
    
ElseIf cap = "print" Then
     cmd_save.Enabled = False
     cmd_back.Enabled = True
     cmd_modify.Enabled = False
     cmd_print.Enabled = True
     cmd_cancel.Enabled = True
     CMD_NEW.Enabled = False

End If
End Sub

Public Sub FLDENABL1(str1 As String)
Dim obj As Object
If str1 = "INI" Then
    'cmb_ACTIVEYN.Text = ""
    MaskEdBox1.Enabled = False
    
    Opt_patient.Enabled = False
    opt_staff.Enabled = False
    
    cmb_reqno.Enabled = False
   
    cmb_enteryby.Enabled = False
    cmb_finalby.Enabled = False
        
    Opt_patient.Value = False
    opt_staff.Value = False

    '*****
    SSTab_MO.Visible = False
    SSTab_MO.TabEnabled(0) = False
    SSTab_MO.TabEnabled(1) = False
    SSTab_MO.TabEnabled(2) = False
    SSTab_MO.TabEnabled(3) = False
    SSTab_MO.TabEnabled(4) = False
    
'    SSTab_MO.TabVisible(0) = False
'    SSTab_MO.TabVisible(1) = False
'    SSTab_MO.TabVisible(2) = False
'    SSTab_MO.TabVisible(3) = False
'    SSTab_MO.TabVisible(4) = False
    
    '***
    ListVirusG408.Enabled = False
ListVirusG409.Enabled = False
ListVirusG406.Enabled = False
ListVirusG410.Enabled = False


ListBacteriaG408.Enabled = False
ListBacteriaG409.Enabled = False
ListBacteriaG406.Enabled = False
ListBacteriaG410.Enabled = False

ListYeastG409.Enabled = False
ListParaG406.Enabled = False
ListAntiG410.Enabled = False
ListAtypicalG410.Enabled = False


ListG407Gpos.Enabled = False
ListG407Gneg.Enabled = False
ListG407Anti.Enabled = False
ListG407Yeast.Enabled = False

Txt_VirusG408.Enabled = False
Txt_VirusG409.Enabled = False
Txt_VirusG406.Enabled = False
Txt_VirusG410.Enabled = False

Txt_BacteriaG408.Enabled = False
Txt_BacteriaG409.Enabled = False
Txt_BacteriaG406.Enabled = False
Txt_BacteriaG410.Enabled = False

Txt_YeastG409.Enabled = False
Txt_ParaG406.Enabled = False
Txt_AntiG410.Enabled = False
Txt_AtypicalG410.Enabled = False

Txt_407Gpos.Enabled = False
Txt_407Gneg.Enabled = False
Txt_407Anti.Enabled = False
Txt_407Yeast.Enabled = False

cmd_AddV408.Enabled = False
cmd_AddV409.Enabled = False
cmd_AddV406.Enabled = False
cmd_AddV410.Enabled = False

cmd_addB408.Enabled = False
cmd_addB409.Enabled = False
cmd_addB406.Enabled = False
cmd_addB410.Enabled = False

cmd_addY409.Enabled = False
cmd_addP406.Enabled = False
cmd_AddAnti410.Enabled = False
cmd_AddAtypical410.Enabled = False

cmd_Add407Gpos.Enabled = False
cmd_Add407Gneg.Enabled = False
cmd_Add407Ganti.Enabled = False
cmd_Add407Gyeast.Enabled = False

SSTab_MO.Visible = True
SSTab_MO.TabEnabled(0) = False
SSTab_MO.TabEnabled(1) = False
SSTab_MO.TabEnabled(2) = False
SSTab_MO.TabEnabled(3) = False
SSTab_MO.TabEnabled(4) = False

'SSTab_MO.TabVisible(0) = False
'SSTab_MO.TabVisible(1) = False
'SSTab_MO.TabVisible(2) = False
'SSTab_MO.TabVisible(3) = False
'SSTab_MO.TabVisible(4) = False
    '***
    
    
    Txt_VirusG408.Enabled = False
    Txt_BacteriaG408.Enabled = False
    
    txt_MOremarks.Text = ""
    
   
    
    Txt_VirusG408.Text = ""
    Txt_BacteriaG408.Text = ""
    

    '*****
    
    MaskEdBox1.Text = ""
    txt_name.Text = ""
    txt_age.Text = ""
    txt_sex.Text = ""
    txt_category.Text = ""
    txt_pathno.Text = ""
    txt_nmaterial.Text = ""
    txt_DMG.Text = ""

    cmb_reqno.clear
    
    
    cmb_finalby.Text = ""
    
    
End If

If str1 = "NEW" Then

    MaskEdBox1.Enabled = False
    Opt_patient.Enabled = True
    opt_staff.Enabled = True
    
    cmb_reqno.Enabled = True
   ' txt_interp.Enabled = True

    cmb_enteryby.Enabled = False
    cmb_finalby.Enabled = True
        
'cmb_ACTIVEYN.Text = ""
    '*****
    '***
    ListVirusG408.Enabled = False
ListVirusG409.Enabled = False
ListVirusG406.Enabled = False
ListVirusG410.Enabled = False


ListBacteriaG408.Enabled = False
ListBacteriaG409.Enabled = False
ListBacteriaG406.Enabled = False
ListBacteriaG410.Enabled = False

ListYeastG409.Enabled = False
ListParaG406.Enabled = False
ListAntiG410.Enabled = False
ListAtypicalG410.Enabled = False


ListG407Gpos.Enabled = False
ListG407Gneg.Enabled = False
ListG407Anti.Enabled = False
ListG407Yeast.Enabled = False

Txt_VirusG408.Enabled = False
Txt_VirusG409.Enabled = False
Txt_VirusG406.Enabled = False
Txt_VirusG410.Enabled = False

Txt_BacteriaG408.Enabled = False
Txt_BacteriaG409.Enabled = False
Txt_BacteriaG406.Enabled = False
Txt_BacteriaG410.Enabled = False

Txt_YeastG409.Enabled = False
Txt_ParaG406.Enabled = False
Txt_AntiG410.Enabled = False
Txt_AtypicalG410.Enabled = False

Txt_407Gpos.Enabled = False
Txt_407Gneg.Enabled = False
Txt_407Anti.Enabled = False
Txt_407Yeast.Enabled = False

cmd_AddV408.Enabled = False
cmd_AddV409.Enabled = False
cmd_AddV406.Enabled = False
cmd_AddV410.Enabled = False

cmd_addB408.Enabled = False
cmd_addB409.Enabled = False
cmd_addB406.Enabled = False
cmd_addB410.Enabled = False

cmd_addY409.Enabled = False
cmd_addP406.Enabled = False
cmd_AddAnti410.Enabled = False
cmd_AddAtypical410.Enabled = False

cmd_Add407Gpos.Enabled = False
cmd_Add407Gneg.Enabled = False
cmd_Add407Ganti.Enabled = False
cmd_Add407Gyeast.Enabled = False

SSTab_MO.Visible = True
SSTab_MO.TabEnabled(0) = False
SSTab_MO.TabEnabled(1) = False
SSTab_MO.TabEnabled(2) = False
SSTab_MO.TabEnabled(3) = False
SSTab_MO.TabEnabled(4) = False

'SSTab_MO.TabVisible(0) = False
'SSTab_MO.TabVisible(1) = False
'SSTab_MO.TabVisible(2) = False
'SSTab_MO.TabVisible(3) = False
'SSTab_MO.TabVisible(4) = False
    '***
    
    '*****
    SSTab_MO.TabEnabled(0) = False
    SSTab_MO.TabEnabled(1) = False
    SSTab_MO.TabEnabled(2) = False
    SSTab_MO.TabEnabled(3) = False
    SSTab_MO.TabEnabled(4) = False
    
'    SSTab_MO.TabVisible(0) = False
'    SSTab_MO.TabVisible(1) = False
'    SSTab_MO.TabVisible(2) = False
'    SSTab_MO.TabVisible(3) = False
'    SSTab_MO.TabVisible(4) = False
    
'    txt_testcat.Text = ""
'    cmb_SampleCondition.Text = ""
'    txt_procedure.Text = ""
'    cmb_Interp.Text = ""
'    txt_interp.Text = ""
'    txt_note.Text = ""
'    txt_comment.Text = ""
    '*****
    
    MaskEdBox1.Text = ""
    txt_name.Text = ""
    txt_age.Text = ""
    txt_sex.Text = ""
    txt_category.Text = ""
    txt_pathno.Text = ""
    txt_nmaterial.Text = ""
    txt_DMG.Text = ""
End If

If str1 = "VIEW" Then
'cmb_ACTIVEYN.Text = ""
    MaskEdBox1.Enabled = False
    
    Opt_patient.Enabled = True
    opt_staff.Enabled = True
    
    cmb_reqno.Enabled = True

    cmb_enteryby.Enabled = False
    cmb_finalby.Enabled = False
        
    
    MaskEdBox1.Text = ""
    txt_name.Text = ""
    txt_age.Text = ""
    txt_sex.Text = ""
    txt_category.Text = ""
    txt_pathno.Text = ""
    txt_nmaterial.Text = ""
    txt_DMG.Text = ""
    '*****
        '***
    ListVirusG408.Enabled = False
ListVirusG409.Enabled = False
ListVirusG406.Enabled = False
ListVirusG410.Enabled = False


ListBacteriaG408.Enabled = False
ListBacteriaG409.Enabled = False
ListBacteriaG406.Enabled = False
ListBacteriaG410.Enabled = False

ListYeastG409.Enabled = False
ListParaG406.Enabled = False
ListAntiG410.Enabled = False
ListAtypicalG410.Enabled = False


ListG407Gpos.Enabled = False
ListG407Gneg.Enabled = False
ListG407Anti.Enabled = False
ListG407Yeast.Enabled = False

Txt_VirusG408.Enabled = False
Txt_VirusG409.Enabled = False
Txt_VirusG406.Enabled = False
Txt_VirusG410.Enabled = False

Txt_BacteriaG408.Enabled = False
Txt_BacteriaG409.Enabled = False
Txt_BacteriaG406.Enabled = False
Txt_BacteriaG410.Enabled = False

Txt_YeastG409.Enabled = False
Txt_ParaG406.Enabled = False
Txt_AntiG410.Enabled = False
Txt_AtypicalG410.Enabled = False

Txt_407Gpos.Enabled = False
Txt_407Gneg.Enabled = False
Txt_407Anti.Enabled = False
Txt_407Yeast.Enabled = False

cmd_AddV408.Enabled = False
cmd_AddV409.Enabled = False
cmd_AddV406.Enabled = False
cmd_AddV410.Enabled = False

cmd_addB408.Enabled = False
cmd_addB409.Enabled = False
cmd_addB406.Enabled = False
cmd_addB410.Enabled = False

cmd_addY409.Enabled = False
cmd_addP406.Enabled = False
cmd_AddAnti410.Enabled = False
cmd_AddAtypical410.Enabled = False

cmd_Add407Gpos.Enabled = False
cmd_Add407Gneg.Enabled = False
cmd_Add407Ganti.Enabled = False
cmd_Add407Gyeast.Enabled = False

SSTab_MO.Visible = True
SSTab_MO.TabEnabled(0) = False
SSTab_MO.TabEnabled(1) = False
SSTab_MO.TabEnabled(2) = False
SSTab_MO.TabEnabled(3) = False
SSTab_MO.TabEnabled(4) = False

'SSTab_MO.TabVisible(0) = False
'SSTab_MO.TabVisible(1) = False
'SSTab_MO.TabVisible(2) = False
'SSTab_MO.TabVisible(3) = False
'SSTab_MO.TabVisible(4) = False
    '***
        
    '*****
    SSTab_MO.TabEnabled(0) = False
            SSTab_MO.TabEnabled(1) = False
            SSTab_MO.TabEnabled(2) = False
            SSTab_MO.TabEnabled(3) = False
            SSTab_MO.TabEnabled(4) = False
            
'            SSTab_MO.TabVisible(0) = False
'            SSTab_MO.TabVisible(1) = False
'            SSTab_MO.TabVisible(2) = False
'            SSTab_MO.TabVisible(3) = False
'            SSTab_MO.TabVisible(4) = False
    
'    txt_testcat.Text = ""
'    cmb_SampleCondition.Text = ""
'    txt_procedure.Text = ""
'    cmb_Interp.Text = ""
'    txt_interp.Text = ""
'    txt_note.Text = ""
'    txt_comment.Text = ""
    '*****
    
End If

If str1 = "MODIFY" Then

    MaskEdBox1.Enabled = False
    
    Opt_patient.Enabled = False
    opt_staff.Enabled = False
    
    cmb_reqno.Enabled = False
    
'    txt_testcat.Enabled = True
'    cmb_SampleCondition.Enabled = True
'    txt_procedure.Enabled = True
'    cmb_Interp.Enabled = True
'    txt_interp.Enabled = True
'    txt_comment.Enabled = True
'    txt_note.Enabled = True

    cmb_enteryby.Enabled = False
    cmb_finalby.Enabled = True
    
    
    
    
    
End If

End Sub





