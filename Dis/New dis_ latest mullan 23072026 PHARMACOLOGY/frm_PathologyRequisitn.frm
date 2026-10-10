VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frm_PathologyRequisitn 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Requisition Form"
   ClientHeight    =   9780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13410
   LinkTopic       =   "Form1"
   ScaleHeight     =   9780
   ScaleWidth      =   13410
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame frm_hematopathology 
      Caption         =   "Endoscopy"
      ForeColor       =   &H000040C0&
      Height          =   3015
      Left            =   120
      TabIndex        =   224
      Top             =   4920
      Visible         =   0   'False
      Width           =   12495
      Begin VB.ComboBox cmb_edprereqno 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   7200
         TabIndex        =   233
         Top             =   2400
         Width           =   4335
      End
      Begin VB.ComboBox cmb_hpprocedure 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1920
         TabIndex        =   232
         Top             =   360
         Width           =   3255
      End
      Begin VB.ComboBox cmb_hpprocedurecode 
         Height          =   315
         Left            =   6720
         TabIndex        =   231
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txt_hpreqno 
         Appearance      =   0  'Flat
         DataField       =   "REQNO"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9360
         TabIndex        =   230
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txt_edsite 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   229
         Top             =   840
         Width           =   6015
      End
      Begin VB.TextBox txt_REQUESTDETAILS 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7200
         TabIndex        =   228
         Top             =   1680
         Width           =   4335
      End
      Begin VB.CommandButton cmd_edindication 
         Caption         =   "&Indication"
         Height          =   495
         Left            =   9840
         TabIndex        =   227
         Top             =   2280
         Width           =   1335
      End
      Begin VB.ListBox list_edts 
         Height          =   1185
         Left            =   1920
         Style           =   1  'Checkbox
         TabIndex        =   226
         Top             =   1560
         Width           =   3135
      End
      Begin VB.ComboBox cmb_edprocedureheading 
         Height          =   315
         Left            =   1920
         TabIndex        =   225
         Top             =   1200
         Width           =   6015
      End
      Begin VB.Label lbl_edprerequisition 
         AutoSize        =   -1  'True
         Caption         =   "Previous Requisition"
         Height          =   195
         Left            =   5400
         TabIndex        =   241
         Top             =   2400
         Width           =   1755
      End
      Begin VB.Label lbl_hpprocedure 
         AutoSize        =   -1  'True
         Caption         =   "Name of Procedure"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   240
         Top             =   480
         Width           =   1650
      End
      Begin VB.Label lbl_hpprocode 
         AutoSize        =   -1  'True
         Caption         =   "Procedure Code"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5280
         TabIndex        =   239
         Top             =   480
         Width           =   1500
      End
      Begin VB.Label lbl_hpreqno1 
         Caption         =   "Requisition No."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7920
         TabIndex        =   238
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lbl_biopsycytology 
         AutoSize        =   -1  'True
         Caption         =   "Tissue Sample"
         Height          =   195
         Left            =   480
         TabIndex        =   237
         Top             =   1680
         Width           =   1245
      End
      Begin VB.Label lbl_siteed 
         AutoSize        =   -1  'True
         Caption         =   "Previous Treatment"
         Height          =   195
         Left            =   120
         TabIndex        =   236
         Top             =   960
         Width           =   1665
      End
      Begin VB.Label lbl_requestremarks 
         AutoSize        =   -1  'True
         Caption         =   "Remarks"
         Height          =   195
         Left            =   6360
         TabIndex        =   235
         Top             =   1800
         Width           =   750
      End
      Begin VB.Label lbl_edprocedureheading 
         AutoSize        =   -1  'True
         Caption         =   "Procedure Heading"
         Height          =   195
         Left            =   120
         TabIndex        =   234
         Top             =   1320
         Width           =   1650
      End
   End
   Begin VB.Frame surgicalPathology 
      Caption         =   "Surgical Pathology"
      Height          =   3015
      Left            =   120
      TabIndex        =   186
      Top             =   4800
      Visible         =   0   'False
      Width           =   12495
      Begin VB.TextBox txt_sppathno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9840
         TabIndex        =   205
         Top             =   480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ComboBox cmb_vpp 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frm_PathologyRequisitn.frx":0000
         Left            =   9840
         List            =   "frm_PathologyRequisitn.frx":0002
         TabIndex        =   203
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox CMB_OPPROCEDURE 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2040
         TabIndex        =   202
         Top             =   1200
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.TextBox txt_metastasis 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2040
         TabIndex        =   201
         Top             =   1560
         Width           =   2175
      End
      Begin VB.ComboBox cmb_frozensec 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frm_PathologyRequisitn.frx":0004
         Left            =   5640
         List            =   "frm_PathologyRequisitn.frx":0006
         TabIndex        =   200
         Top             =   1920
         Width           =   1575
      End
      Begin VB.ComboBox cmb_preinvest 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6480
         TabIndex        =   199
         Top             =   2520
         Width           =   5175
      End
      Begin VB.ComboBox cmb_laterality 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5640
         Sorted          =   -1  'True
         TabIndex        =   198
         Top             =   1560
         Width           =   1575
      End
      Begin VB.ComboBox cmb_sites 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   197
         Top             =   720
         Width           =   2775
      End
      Begin VB.ComboBox cmb_laterality1 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6480
         Sorted          =   -1  'True
         TabIndex        =   196
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox cmb_procedurecode 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frm_PathologyRequisitn.frx":0008
         Left            =   9840
         List            =   "frm_PathologyRequisitn.frx":000A
         TabIndex        =   195
         Top             =   120
         Width           =   1815
      End
      Begin VB.ComboBox cmb_procedure 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frm_PathologyRequisitn.frx":000C
         Left            =   2040
         List            =   "frm_PathologyRequisitn.frx":000E
         TabIndex        =   194
         Top             =   240
         Width           =   6195
      End
      Begin VB.ComboBox cmb_outblock 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2040
         TabIndex        =   193
         Top             =   1920
         Width           =   1575
      End
      Begin VB.ComboBox cmb_outslide 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9240
         TabIndex        =   192
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox txt_outspecimen 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9240
         TabIndex        =   191
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txt_nmeterial 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8880
         MultiLine       =   -1  'True
         TabIndex        =   190
         ToolTipText     =   "Double Click to open  text document"
         Top             =   1200
         Width           =   2775
      End
      Begin VB.ComboBox cmb_spfnac 
         Height          =   315
         ItemData        =   "frm_PathologyRequisitn.frx":0010
         Left            =   2040
         List            =   "frm_PathologyRequisitn.frx":0012
         TabIndex        =   189
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txt_spreqno 
         Appearance      =   0  'Flat
         DataField       =   "REQNO"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   9840
         TabIndex        =   188
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox TXT_OPPROCEDURE 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         TabIndex        =   187
         Top             =   1200
         Width           =   5175
      End
      Begin MSMask.MaskEdBox MSK_FROSECDT 
         Height          =   375
         Left            =   4800
         TabIndex        =   204
         Top             =   2400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lbl_sppathno 
         AutoSize        =   -1  'True
         Caption         =   "NIP"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9360
         TabIndex        =   223
         Top             =   960
         Width           =   330
      End
      Begin VB.Label lbl_spreqno1 
         AutoSize        =   -1  'True
         Caption         =   "Requisition No."
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   8400
         TabIndex        =   222
         Top             =   600
         Width           =   1320
      End
      Begin VB.Label lbl_procedure 
         AutoSize        =   -1  'True
         Caption         =   "Name of Procedure"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   221
         Top             =   360
         Width           =   1650
      End
      Begin VB.Label lbl_procedurecode 
         AutoSize        =   -1  'True
         Caption         =   "Pro. Code"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8760
         TabIndex        =   220
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lbl_Preinvest 
         AutoSize        =   -1  'True
         Caption         =   "Prev Reqn No/Path No"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7440
         TabIndex        =   219
         Top             =   2280
         Width           =   2235
      End
      Begin VB.Label lbl_frdate 
         AutoSize        =   -1  'True
         Caption         =   "FrozenSection Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3000
         TabIndex        =   218
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label lbl_frozensec 
         AutoSize        =   -1  'True
         Caption         =   "Frozen Section"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4305
         TabIndex        =   217
         Top             =   2040
         Width           =   1290
      End
      Begin VB.Label lbl_outsideslide 
         AutoSize        =   -1  'True
         Caption         =   "Outside Slide/s"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7875
         TabIndex        =   216
         Top             =   2040
         Width           =   1320
      End
      Begin VB.Label lbl_outsideblock 
         AutoSize        =   -1  'True
         Caption         =   "Outside Block/s"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   615
         TabIndex        =   215
         Top             =   2040
         Width           =   1380
      End
      Begin VB.Label lbl_outside 
         AutoSize        =   -1  'True
         Caption         =   "Outside Specimen/s"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7470
         TabIndex        =   214
         Top             =   1680
         Width           =   1725
      End
      Begin VB.Label lbl_lit 
         AutoSize        =   -1  'True
         Caption         =   "Laterality"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4725
         TabIndex        =   213
         Top             =   1560
         Width           =   795
      End
      Begin VB.Label lbl_meta 
         Caption         =   "Metastasis"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1020
         TabIndex        =   212
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lbl_Opprocedure 
         AutoSize        =   -1  'True
         Caption         =   "Operative Procedure/s"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   45
         TabIndex        =   211
         Top             =   1200
         Width           =   1950
      End
      Begin VB.Label lbl_laterality 
         AutoSize        =   -1  'True
         Caption         =   "Laterality"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5640
         TabIndex        =   210
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lbl_site 
         AutoSize        =   -1  'True
         Caption         =   "Site/s"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1200
         TabIndex        =   209
         Top             =   720
         Width           =   525
      End
      Begin VB.Label LBL_NMETERIAL 
         Caption         =   "Nature of Material sent"
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   7320
         TabIndex        =   208
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label LBL_SPFNAC 
         AutoSize        =   -1  'True
         Caption         =   "F.N.A.C"
         Height          =   195
         Left            =   1320
         TabIndex        =   207
         Top             =   2520
         Width           =   675
      End
      Begin VB.Label lbl_sppathno1 
         Caption         =   "Path No."
         Height          =   255
         Left            =   9000
         TabIndex        =   206
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame Microbiology 
      Caption         =   "Microbiology"
      ForeColor       =   &H000000C0&
      Height          =   2895
      Left            =   120
      TabIndex        =   164
      Top             =   4800
      Visible         =   0   'False
      Width           =   11775
      Begin VB.ComboBox CMB_ICU 
         Height          =   315
         ItemData        =   "frm_PathologyRequisitn.frx":0014
         Left            =   5040
         List            =   "frm_PathologyRequisitn.frx":001E
         TabIndex        =   175
         Text            =   "No"
         Top             =   2520
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txt_mbreqno 
         Appearance      =   0  'Flat
         DataField       =   "REQNO"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9480
         TabIndex        =   174
         Top             =   360
         Width           =   2175
      End
      Begin VB.ComboBox cmb_mbprevunv 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7800
         TabIndex        =   173
         Top             =   2520
         Width           =   3855
      End
      Begin VB.ComboBox cmb_mboutslide 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1920
         TabIndex        =   172
         Top             =   2520
         Width           =   2055
      End
      Begin VB.TextBox txt_mbspecimen 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   6240
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   171
         Top             =   1320
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.ListBox list_proc 
         ForeColor       =   &H00000000&
         Height          =   960
         ItemData        =   "frm_PathologyRequisitn.frx":002B
         Left            =   1935
         List            =   "frm_PathologyRequisitn.frx":002D
         Style           =   1  'Checkbox
         TabIndex        =   170
         Top             =   240
         Width           =   4095
      End
      Begin VB.ListBox List_code 
         ForeColor       =   &H00000000&
         Height          =   960
         ItemData        =   "frm_PathologyRequisitn.frx":002F
         Left            =   7080
         List            =   "frm_PathologyRequisitn.frx":0031
         Style           =   1  'Checkbox
         TabIndex        =   169
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txt_ongoinganti 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   1920
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   168
         Top             =   1800
         Width           =   9735
      End
      Begin VB.TextBox TXT_MBPATHNO 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9480
         TabIndex        =   167
         Top             =   720
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox CMB_MBNIP 
         Height          =   315
         Left            =   9480
         TabIndex        =   166
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox cmb_specimen 
         Height          =   315
         Left            =   1920
         TabIndex        =   165
         Top             =   1320
         Width           =   4095
      End
      Begin VB.Label lbl_mbreqno 
         AutoSize        =   -1  'True
         Caption         =   "Requisition No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8160
         TabIndex        =   185
         Top             =   360
         Width           =   1320
      End
      Begin VB.Label lbl_mbpreinv 
         AutoSize        =   -1  'True
         Caption         =   "Prev. Investigated"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6120
         TabIndex        =   184
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label lbl_mboutslides 
         AutoSize        =   -1  'True
         Caption         =   "Outside Slides"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   183
         Top             =   2640
         Width           =   1230
      End
      Begin VB.Label lbl_mbspecimendesc 
         AutoSize        =   -1  'True
         Caption         =   "Specimen "
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   960
         TabIndex        =   182
         Top             =   1440
         Width           =   900
      End
      Begin VB.Label lbl_mbongant 
         AutoSize        =   -1  'True
         Caption         =   "Ongoing Antimicrobial"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   0
         TabIndex        =   181
         Top             =   2040
         Width           =   1860
      End
      Begin VB.Label lbl_mbprocode 
         Caption         =   "Procedure Code"
         ForeColor       =   &H00000000&
         Height          =   555
         Left            =   6120
         TabIndex        =   180
         Top             =   480
         Width           =   900
      End
      Begin VB.Label lbl_mbprocedure 
         AutoSize        =   -1  'True
         Caption         =   "Name of Procedure"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   179
         Top             =   720
         Width           =   1650
      End
      Begin VB.Label LBL_MBMICRONO 
         AutoSize        =   -1  'True
         Caption         =   "Micro No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8640
         TabIndex        =   178
         Top             =   720
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label LBL_MBNIP 
         AutoSize        =   -1  'True
         Caption         =   "NIP"
         Height          =   195
         Left            =   9120
         TabIndex        =   177
         Top             =   1080
         Width           =   330
      End
      Begin VB.Label LBL_ICU 
         AutoSize        =   -1  'True
         Caption         =   "ICU"
         Height          =   195
         Left            =   4560
         TabIndex        =   176
         Top             =   2640
         Visible         =   0   'False
         Width           =   330
      End
   End
   Begin VB.Frame immunohistochemistry 
      Caption         =   "Immunihistochemistry"
      ForeColor       =   &H000000C0&
      Height          =   3015
      Left            =   120
      TabIndex        =   143
      Top             =   4800
      Visible         =   0   'False
      Width           =   11775
      Begin VB.ComboBox cmb_Ihsoutsideblock 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2640
         TabIndex        =   153
         Top             =   1680
         Width           =   2055
      End
      Begin VB.ComboBox cmb_Ihsoutslide 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9720
         TabIndex        =   152
         Top             =   1680
         Width           =   1815
      End
      Begin VB.ComboBox cmb_Ihspreinv 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2640
         TabIndex        =   151
         Top             =   2280
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.ComboBox cmb_ihprocedure 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2640
         TabIndex        =   150
         Top             =   480
         Width           =   2175
      End
      Begin VB.ComboBox cmb_ihprocode 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6600
         TabIndex        =   149
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txt_sppathreqno 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2640
         TabIndex        =   148
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox txt_ihreqno 
         Appearance      =   0  'Flat
         DataField       =   "REQNO"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9600
         TabIndex        =   147
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txt_ihpathno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9600
         TabIndex        =   146
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txt_ihmeterial 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8280
         TabIndex        =   145
         Top             =   1200
         Width           =   3255
      End
      Begin VB.ComboBox CMB_IHNIP 
         Height          =   315
         Left            =   9720
         TabIndex        =   144
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lbl_hisblock 
         AutoSize        =   -1  'True
         Caption         =   "Outside Block"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1320
         TabIndex        =   163
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label lbl_hisslide 
         AutoSize        =   -1  'True
         Caption         =   "Outside Slides"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8280
         TabIndex        =   162
         Top             =   1680
         Width           =   1230
      End
      Begin VB.Label lbl_hispreins 
         AutoSize        =   -1  'True
         Caption         =   "Previously Investigated"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   480
         TabIndex        =   161
         Top             =   2280
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label lbl_ihprocedure 
         AutoSize        =   -1  'True
         Caption         =   "Number of  Procedure"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   160
         Top             =   600
         Width           =   1875
      End
      Begin VB.Label lbl_ihprocdoe 
         AutoSize        =   -1  'True
         Caption         =   "Procedure Code"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5160
         TabIndex        =   159
         Top             =   600
         Width           =   1380
      End
      Begin VB.Label lbl_spreqno 
         AutoSize        =   -1  'True
         Caption         =   "Surgical Path Requisition No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   158
         Top             =   1080
         Width           =   2520
      End
      Begin VB.Label lbl_ihreqno 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Requisition No."
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   8280
         TabIndex        =   157
         Top             =   480
         Width           =   1320
      End
      Begin VB.Label lbl_ihpathno 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Path No."
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   8880
         TabIndex        =   156
         Top             =   840
         Width           =   765
      End
      Begin VB.Label lbl_ihnmeterial 
         AutoSize        =   -1  'True
         Caption         =   "Nature of Meterial"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   6600
         TabIndex        =   155
         Top             =   1320
         Width           =   1545
      End
      Begin VB.Label LBL_IHNIP 
         AutoSize        =   -1  'True
         Caption         =   "NIP"
         Height          =   195
         Left            =   9120
         TabIndex        =   154
         Top             =   2280
         Width           =   330
      End
   End
   Begin VB.Frame CYTOLOGY 
      Caption         =   "Gynaec/NonGynaec"
      ForeColor       =   &H000000C0&
      Height          =   3015
      Left            =   120
      TabIndex        =   112
      Top             =   4800
      Visible         =   0   'False
      Width           =   11775
      Begin VB.TextBox cmb_cyoutspecimen 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6000
         TabIndex        =   126
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txt_cyreqno1 
         Appearance      =   0  'Flat
         DataField       =   "REQNO"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9600
         TabIndex        =   125
         Top             =   390
         Width           =   2055
      End
      Begin VB.ComboBox cmb_cpreinv 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7560
         TabIndex        =   124
         Top             =   2400
         Width           =   4095
      End
      Begin VB.TextBox txt_trtdetail 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9600
         TabIndex        =   123
         Top             =   840
         Width           =   2055
      End
      Begin VB.ComboBox cmb_coutslides 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9480
         TabIndex        =   122
         Top             =   1680
         Width           =   2175
      End
      Begin VB.ComboBox cmb_claterality 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6000
         Sorted          =   -1  'True
         TabIndex        =   121
         Top             =   810
         Width           =   1935
      End
      Begin VB.TextBox txt_cmetastasis 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6000
         TabIndex        =   120
         Top             =   1320
         Width           =   1935
      End
      Begin VB.ComboBox cmb_cnoslides 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1920
         TabIndex        =   119
         Top             =   1680
         Width           =   1350
      End
      Begin VB.TextBox txt_cnatmeterial 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1920
         TabIndex        =   118
         Top             =   1320
         Width           =   2775
      End
      Begin VB.ComboBox cbo_proccode 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6360
         TabIndex        =   117
         Top             =   390
         Width           =   1455
      End
      Begin VB.TextBox txt_scopyfind 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1920
         TabIndex        =   116
         Top             =   2160
         Width           =   2775
      End
      Begin VB.ComboBox CMB_CGPROCEDURE 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1920
         TabIndex        =   115
         Top             =   360
         Width           =   2775
      End
      Begin VB.ComboBox cmb_cnsite 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   114
         Top             =   810
         Width           =   2775
      End
      Begin VB.ComboBox CMB_CGNIP 
         Height          =   315
         Left            =   9600
         TabIndex        =   113
         Top             =   1200
         Width           =   1575
      End
      Begin MSMask.MaskEdBox MSK_LMP 
         Height          =   375
         Left            =   6000
         TabIndex        =   127
         Top             =   2280
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   0
         PromptInclude   =   0   'False
         HideSelection   =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lbl_cyreqno 
         AutoSize        =   -1  'True
         Caption         =   "Requisition No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8280
         TabIndex        =   142
         Top             =   480
         Width           =   1320
      End
      Begin VB.Label lbl_lmp 
         AutoSize        =   -1  'True
         Caption         =   "LMP"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5445
         TabIndex        =   141
         Top             =   2400
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label lbl_cpreinvest 
         AutoSize        =   -1  'True
         Caption         =   "Previously Investigated"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7440
         TabIndex        =   140
         Top             =   2160
         Width           =   1995
      End
      Begin VB.Label lbl_ctrtdetail 
         AutoSize        =   -1  'True
         Caption         =   "Treatment Details"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   8085
         TabIndex        =   139
         Top             =   930
         Width           =   1515
      End
      Begin VB.Label lbl_outslide 
         AutoSize        =   -1  'True
         Caption         =   "Outside Slides"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8205
         TabIndex        =   138
         Top             =   1800
         Width           =   1230
      End
      Begin VB.Label lbl_claterality 
         AutoSize        =   -1  'True
         Caption         =   "Laterality"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5040
         TabIndex        =   137
         Top             =   930
         Width           =   795
      End
      Begin VB.Label lbl_cmetastasis 
         Caption         =   "Metastasis"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4980
         TabIndex        =   136
         Top             =   1350
         Width           =   975
      End
      Begin VB.Label lbl_coutspecimen 
         AutoSize        =   -1  'True
         Caption         =   "Outside Specimen"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   4290
         TabIndex        =   135
         Top             =   1920
         Width           =   1545
      End
      Begin VB.Label lbl_cnoslides 
         AutoSize        =   -1  'True
         Caption         =   "No of Slides"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   705
         TabIndex        =   134
         Top             =   1800
         Width           =   1050
      End
      Begin VB.Label lbl_cnatumaterial 
         AutoSize        =   -1  'True
         Caption         =   "Nature of Material"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   210
         TabIndex        =   133
         Top             =   1410
         Width           =   1545
      End
      Begin VB.Label lbl_csiteofspec 
         AutoSize        =   -1  'True
         Caption         =   "Site"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1410
         TabIndex        =   132
         Top             =   930
         Width           =   345
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Name of procedure"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   131
         Top             =   360
         Width           =   1635
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Procedure Code"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4920
         TabIndex        =   130
         Top             =   480
         Width           =   1380
      End
      Begin VB.Label lbl_scopyfind 
         AutoSize        =   -1  'True
         Caption         =   "OT Procedure No."
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   240
         TabIndex        =   129
         Top             =   2160
         Width           =   1560
      End
      Begin VB.Label lbl_cgnip 
         AutoSize        =   -1  'True
         Caption         =   "NIP"
         Height          =   195
         Left            =   9120
         TabIndex        =   128
         Top             =   1320
         Width           =   330
      End
   End
   Begin VB.Frame bceltm 
      Height          =   3015
      Left            =   120
      TabIndex        =   105
      Top             =   4800
      Visible         =   0   'False
      Width           =   11775
      Begin VB.TextBox txt_bcreqno 
         Appearance      =   0  'Flat
         DataField       =   "REQNO"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   9360
         TabIndex        =   108
         Top             =   360
         Width           =   2175
      End
      Begin VB.ComboBox cmb_cyprocode 
         Height          =   315
         Left            =   9795
         TabIndex        =   107
         Top             =   1080
         Width           =   1515
      End
      Begin VB.ComboBox cmb_cyprocedure 
         Height          =   315
         Left            =   585
         TabIndex        =   106
         Top             =   1095
         Width           =   7845
      End
      Begin VB.Label lbl_bcreqno 
         AutoSize        =   -1  'True
         Caption         =   "Requisition No."
         Height          =   195
         Left            =   7920
         TabIndex        =   111
         Top             =   480
         Width           =   1320
      End
      Begin VB.Label Label11 
         Caption         =   "Test Code"
         Height          =   270
         Left            =   8820
         TabIndex        =   110
         Top             =   1155
         Width           =   930
      End
      Begin VB.Label Label10 
         Caption         =   "Test"
         Height          =   255
         Left            =   120
         TabIndex        =   109
         Top             =   1125
         Width           =   420
      End
   End
   Begin VB.Frame fnac 
      Caption         =   "FNAC"
      ForeColor       =   &H000000C0&
      Height          =   2895
      Left            =   120
      TabIndex        =   76
      Top             =   4800
      Visible         =   0   'False
      Width           =   11775
      Begin VB.TextBox txt_fnnmeterial 
         Appearance      =   0  'Flat
         Height          =   645
         Left            =   9600
         TabIndex        =   90
         Top             =   1080
         Width           =   2055
      End
      Begin VB.ComboBox cmb_outsideslidesfn 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9600
         TabIndex        =   89
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox txt_fnmetastasis 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         TabIndex        =   88
         Top             =   1320
         Width           =   2895
      End
      Begin VB.ComboBox cmb_fnlaterality1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6000
         Sorted          =   -1  'True
         TabIndex        =   87
         Top             =   1320
         Width           =   1815
      End
      Begin VB.ComboBox cmb_fnlaterality 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6000
         Sorted          =   -1  'True
         TabIndex        =   86
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox cmb_fnpreinvest 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   7800
         TabIndex        =   85
         Top             =   2400
         Width           =   3855
      End
      Begin VB.TextBox txt_FNreqno 
         Appearance      =   0  'Flat
         DataField       =   "REQNO"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   9840
         TabIndex        =   84
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox CMB_FNPROCEDURE 
         Height          =   315
         Left            =   2040
         TabIndex        =   83
         Top             =   360
         Width           =   2895
      End
      Begin VB.ComboBox cmb_fnprocode 
         Height          =   315
         Left            =   6600
         TabIndex        =   82
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox cmb_fnac 
         Height          =   315
         Left            =   2040
         TabIndex        =   81
         Top             =   2160
         Width           =   2895
      End
      Begin VB.ComboBox cmb_fnsite 
         Height          =   315
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   80
         Top             =   840
         Width           =   2895
      End
      Begin VB.ComboBox CMB_FANIP 
         Height          =   315
         Left            =   9600
         TabIndex        =   79
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox TXT_FNSLIDEDRY 
         Height          =   315
         Left            =   2040
         TabIndex        =   78
         Top             =   1680
         Width           =   1815
      End
      Begin VB.ComboBox TXT_FNOSLIDEWET 
         Height          =   315
         Left            =   6000
         TabIndex        =   77
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label lbl_sitefnac 
         AutoSize        =   -1  'True
         Caption         =   "Site of FNAC"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   840
         TabIndex        =   104
         Top             =   840
         Width           =   1110
      End
      Begin VB.Label LBL_CYLATERALITY 
         AutoSize        =   -1  'True
         Caption         =   "Laterality"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5160
         TabIndex        =   103
         Top             =   840
         Width           =   795
      End
      Begin VB.Label lbl_nmaterial 
         AutoSize        =   -1  'True
         Caption         =   "Nature of Material"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   7920
         TabIndex        =   102
         Top             =   1440
         Width           =   1545
      End
      Begin VB.Label lbl_nofslidesdry 
         Caption         =   "No of Slides Dry"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   600
         TabIndex        =   101
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lbl_noofswet 
         AutoSize        =   -1  'True
         Caption         =   "No. of Slides Wet"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4440
         TabIndex        =   100
         Top             =   1800
         Width           =   1515
      End
      Begin VB.Label lbl_FANIP 
         AutoSize        =   -1  'True
         Caption         =   "NIP"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9120
         TabIndex        =   99
         Top             =   840
         Width           =   330
      End
      Begin VB.Label lbl_fnac 
         AutoSize        =   -1  'True
         Caption         =   "FNAC Performed by"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   98
         Top             =   2160
         Width           =   1665
      End
      Begin VB.Label lbl_fnmetastasis 
         AutoSize        =   -1  'True
         Caption         =   "Metastasis"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1080
         TabIndex        =   97
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label lbl1_laterality 
         AutoSize        =   -1  'True
         Caption         =   "Laterality"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5160
         TabIndex        =   96
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label lbl_outsideslides 
         AutoSize        =   -1  'True
         Caption         =   "Outside Slides"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8280
         TabIndex        =   95
         Top             =   1920
         Width           =   1230
      End
      Begin VB.Label lbl_proinv1 
         AutoSize        =   -1  'True
         Caption         =   "Previously Investigated"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5760
         TabIndex        =   94
         Top             =   2520
         Width           =   1995
      End
      Begin VB.Label lbl_fnreqno 
         AutoSize        =   -1  'True
         Caption         =   "Requisition No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8400
         TabIndex        =   93
         Top             =   480
         Width           =   1320
      End
      Begin VB.Label LBL_FNPROCEDURE 
         AutoSize        =   -1  'True
         Caption         =   "Name of Procedure"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   92
         Top             =   360
         Width           =   1650
      End
      Begin VB.Label lbl_fnprocode 
         AutoSize        =   -1  'True
         Caption         =   "Procedure Code"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5160
         TabIndex        =   91
         Top             =   360
         Width           =   1380
      End
   End
   Begin VB.Frame common 
      ForeColor       =   &H00C00000&
      Height          =   3855
      Left            =   120
      TabIndex        =   18
      Top             =   960
      Width           =   12495
      Begin VB.ComboBox CMB_TEXT 
         Height          =   315
         Left            =   5760
         TabIndex        =   49
         Top             =   3000
         Width           =   1935
      End
      Begin VB.ComboBox cmb_dept 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2520
         Sorted          =   -1  'True
         TabIndex        =   48
         Top             =   3000
         Width           =   3135
      End
      Begin VB.TextBox txt_dept 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10680
         TabIndex        =   47
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox txt_name 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   720
         Width           =   5895
      End
      Begin VB.TextBox txt_sex 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2520
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txt_Age 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4680
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txt_status 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10800
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txt_catgory 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6840
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txt_pridocser 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2520
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox txt_ward 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6840
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox txt_bedno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10800
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txt_modcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10680
         TabIndex        =   37
         Top             =   3120
         Width           =   1455
      End
      Begin VB.ComboBox cmb_modality 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frm_PathologyRequisitn.frx":0033
         Left            =   2520
         List            =   "frm_PathologyRequisitn.frx":0035
         TabIndex        =   36
         Top             =   3480
         Width           =   3735
      End
      Begin VB.TextBox txt_clinhist 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2520
         TabIndex        =   35
         Top             =   2280
         Width           =   5895
      End
      Begin VB.ComboBox CMB_DURILLMONTH 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   11520
         TabIndex        =   34
         Top             =   2280
         Width           =   615
      End
      Begin VB.ComboBox cmb_reqno 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5295
         TabIndex        =   33
         Top             =   240
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txt_pathno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10800
         TabIndex        =   32
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cmb_prodis 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2520
         Sorted          =   -1  'True
         TabIndex        =   31
         Top             =   2640
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.ComboBox cmb_refdoc 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2520
         Sorted          =   -1  'True
         TabIndex        =   30
         Top             =   1920
         Width           =   2175
      End
      Begin VB.ComboBox cmb_reftodoc 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6840
         Sorted          =   -1  'True
         TabIndex        =   29
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txt_year 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10680
         TabIndex        =   28
         Top             =   2280
         Width           =   735
      End
      Begin VB.Frame emppatient 
         Height          =   1095
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   1455
         Begin VB.OptionButton opt_patient 
            Caption         =   "Patient"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton opt_staff 
            Caption         =   "Staff"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.ComboBox cmb_dependents 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   10800
         TabIndex        =   24
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TEXT1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7200
         TabIndex        =   23
         Top             =   3480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox CMB_ICU1 
         Height          =   315
         Left            =   10680
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   3480
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8040
         TabIndex        =   21
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "&EMR"
         Height          =   375
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txt_indctinvstg 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   19
         Top             =   2640
         Width           =   3135
      End
      Begin MSMask.MaskEdBox Msk_caseno 
         Height          =   375
         Left            =   2400
         TabIndex        =   46
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   0
         PromptInclude   =   0   'False
         MaxLength       =   8
         Mask            =   "??>/#####"
         PromptChar      =   "_"
      End
      Begin VB.Label lbl_icureq 
         Caption         =   "ICU Requisition"
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
         Left            =   9000
         TabIndex        =   75
         Top             =   3480
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Label lbl_rwd 
         Caption         =   "RECOVERY WARD"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   9960
         TabIndex        =   74
         Top             =   3480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lbl_dept 
         AutoSize        =   -1  'True
         Caption         =   "Requisition Placed to"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   480
         TabIndex        =   73
         Top             =   3120
         Width           =   1830
      End
      Begin VB.Label lbl_deptcode 
         AutoSize        =   -1  'True
         Caption         =   "Dept. Code"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9000
         TabIndex        =   72
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label lbl_caseno 
         AutoSize        =   -1  'True
         Caption         =   "Case No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1515
         TabIndex        =   71
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lbl_name 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1815
         TabIndex        =   70
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl_sex 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1980
         TabIndex        =   69
         Top             =   1320
         Width           =   330
      End
      Begin VB.Label lbl_pridocser 
         AutoSize        =   -1  'True
         Caption         =   "Primary Doctor/Service"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   315
         TabIndex        =   68
         Top             =   1680
         Width           =   1995
      End
      Begin VB.Label lbl_age 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4200
         TabIndex        =   67
         Top             =   1320
         Width           =   345
      End
      Begin VB.Label lbl_status 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9720
         TabIndex        =   66
         Top             =   1200
         Width           =   555
      End
      Begin VB.Label lbl_catgory 
         AutoSize        =   -1  'True
         Caption         =   "Category"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5970
         TabIndex        =   65
         Top             =   1320
         Width           =   765
      End
      Begin VB.Label lbl_ward 
         AutoSize        =   -1  'True
         Caption         =   "Ward"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6270
         TabIndex        =   64
         Top             =   1680
         Width           =   465
      End
      Begin VB.Label lbl_bedno 
         AutoSize        =   -1  'True
         Caption         =   "Bed No"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9630
         TabIndex        =   63
         Top             =   1560
         Width           =   645
      End
      Begin VB.Label lbl_modcode 
         AutoSize        =   -1  'True
         Caption         =   "Modality Code"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8760
         TabIndex        =   62
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label lbl_modality 
         AutoSize        =   -1  'True
         Caption         =   "Modality"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1590
         TabIndex        =   61
         Top             =   3600
         Width           =   720
      End
      Begin VB.Label lbl_refdoc 
         AutoSize        =   -1  'True
         Caption         =   "Referring Doctor/Service"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   150
         TabIndex        =   60
         Top             =   2040
         Width           =   2160
      End
      Begin VB.Label lbl_reftodoc 
         AutoSize        =   -1  'True
         Caption         =   "Referred to Doctor"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5160
         TabIndex        =   59
         Top             =   2040
         Width           =   1605
      End
      Begin VB.Label lbl_clinhist 
         AutoSize        =   -1  'True
         Caption         =   "Relevant Clinical History"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   255
         TabIndex        =   58
         Top             =   2400
         Width           =   2100
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Duration of Illness"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9000
         TabIndex        =   57
         Top             =   2280
         Width           =   1440
      End
      Begin VB.Label lbl_provdis 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Indication For Investigations"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   30
         TabIndex        =   56
         Top             =   2760
         Width           =   2445
      End
      Begin VB.Label lbl_reqno 
         Caption         =   "Requisition No."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3960
         TabIndex        =   55
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lbl_pathno 
         AutoSize        =   -1  'True
         Caption         =   "Path No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9840
         TabIndex        =   54
         Top             =   240
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_year 
         AutoSize        =   -1  'True
         Caption         =   "Years      Months"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   10680
         TabIndex        =   53
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label LBL_DEPENDENT 
         AutoSize        =   -1  'True
         Caption         =   "Dependents"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9480
         TabIndex        =   52
         Top             =   720
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "DTM No."
         Height          =   255
         Left            =   6360
         TabIndex        =   51
         Top             =   3480
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Outside Sample"
         Height          =   255
         Left            =   8280
         TabIndex        =   50
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.TextBox txt_date 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   10680
      TabIndex        =   17
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txt_time 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   10665
      TabIndex        =   16
      Top             =   540
      Width           =   1215
   End
   Begin VB.Frame buttons 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   7800
      Width           =   12495
      Begin VB.CommandButton CMD_ADDNEW 
         Caption         =   "ADDNEW"
         Height          =   375
         Left            =   1440
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmd_new 
         Caption         =   "&New"
         Default         =   -1  'True
         Height          =   375
         Left            =   240
         MouseIcon       =   "frm_PathologyRequisitn.frx":0037
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_view 
         Caption         =   "&View"
         Height          =   375
         Left            =   2400
         MouseIcon       =   "frm_PathologyRequisitn.frx":0341
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmd_modify 
         Caption         =   "&Modify"
         Height          =   375
         Left            =   3360
         MouseIcon       =   "frm_PathologyRequisitn.frx":064B
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_save 
         Caption         =   "&Save"
         Height          =   375
         Left            =   4560
         MouseIcon       =   "frm_PathologyRequisitn.frx":0955
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   5280
         MouseIcon       =   "frm_PathologyRequisitn.frx":0C5F
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "&Print"
         Height          =   375
         Left            =   6480
         MouseIcon       =   "frm_PathologyRequisitn.frx":0F69
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_back 
         Caption         =   "&Back"
         Height          =   375
         Left            =   9960
         MouseIcon       =   "frm_PathologyRequisitn.frx":1273
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_repeat 
         Caption         =   "&Repeat"
         Height          =   375
         Left            =   8880
         MouseIcon       =   "frm_PathologyRequisitn.frx":157D
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CMB_MEMO 
         Caption         =   "&Memo"
         Height          =   375
         Left            =   7680
         MouseIcon       =   "frm_PathologyRequisitn.frx":1887
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H80000010&
         Caption         =   "Print Barcode"
         Enabled         =   0   'False
         Height          =   375
         Left            =   11160
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   10320
      TabIndex        =   3
      Top             =   8880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6960
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.ComboBox CMBIRB 
      Height          =   315
      Left            =   8400
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
   Begin MSMask.MaskEdBox MSK1 
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSWinsockLib.Winsock pathwin 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lbl_date 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   10080
      TabIndex        =   249
      Top             =   240
      Width           =   540
   End
   Begin VB.Label lbl_time 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   10080
      TabIndex        =   248
      Top             =   600
      Width           =   540
   End
   Begin VB.Label lbl_req 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "REQUISITION"
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
      Left            =   4800
      TabIndex        =   247
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   495
      Left            =   5520
      TabIndex        =   246
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   5520
      TabIndex        =   245
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   5520
      TabIndex        =   244
      Top             =   4080
      Width           =   1215
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   555
      TabIndex        =   243
      Top             =   0
      Width           =   8595
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "PROTOCOL PATIENT"
      Enabled         =   0   'False
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   120
      TabIndex        =   242
      Top             =   720
      Visible         =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "frm_PathologyRequisitn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SHW_REQNO As String

'molecular pathology
Dim BMEMOEN As Boolean

Dim SELDREQMPQR As rdoQuery
Dim SELDREQMPRS As rdoResultset
Dim seLdreqrsmpflag As Boolean
Dim INSDMOLPROC_QR As rdoQuery
Dim INSDMOLREQ_QR As rdoQuery

Dim keyflag As Boolean
Dim otlnk As Boolean

Dim SELELEMP_QR As rdoQuery
Dim SELELEMP_RS As rdoResultset
Dim SELELEMPFLAG As Boolean

Dim SELSCOPYFNDQR As rdoQuery
Dim SELSCOPYFNDRS As rdoResultset
Dim SELSCOPYFNDFLAG As Boolean

'selecting provisional diagnosis
Dim selprvdgs_qr As rdoQuery
Dim selprvdgs_rs As rdoResultset
Dim selprvdgs As Boolean

'CYTOGENETIC
Dim INSRREQUISITNCY_QR As rdoQuery
Dim SELRREQUISITNCY_QR As rdoQuery
Dim SELRREQUISITNCY_RS As rdoResultset
Dim UPDRREQUISITNCY_qr As rdoQuery
Dim INSDREQUISITNCY_QR As rdoQuery

Dim PRFREE_QR As rdoQuery
Dim PRFREE_RS As rdoResultset
Dim PreFreeCode As String
Dim PRFREE_FLG As Boolean

Dim INSCYTOGCLDTACY_QR As rdoQuery
Dim UPDCYTOGCLDTACY_QR As rdoQuery
Dim SELCYTOGCLDTACY_QR As rdoQuery
Dim SELCYTOGCLDTACY_RS As rdoResultset
Dim SELCYTOGCLDTACY_FLG As Boolean

Dim SELMA101CY_QR As rdoQuery
Dim SELMA101CY_RS As rdoResultset
Dim SELMA101CY_FLAG As Boolean

Dim selreqmoda_qr As rdoQuery
Dim selreqmoda_RS As rdoResultset

Dim selreqnno_qr As rdoQuery
Dim selreqnno_rs As rdoResultset

Dim SELRINVESTCY_QR As rdoQuery
Dim SELRINVESTCY_RS As rdoResultset
Dim SELRINVESTCY_FLAG As Boolean



'CYTOGENETIC
Dim I As Integer
Dim I1 As Integer
Dim CASENO As String
Dim AUTOREQNO As Variant
Dim DAUTOREQNO As Variant
Dim AUTOPATHNO As Variant
Dim DAUTOPATHNO As Variant
Dim PATYPE As String
Dim viewflag As Boolean
Dim NEWFLAG As Boolean
Dim PRINTNEW As Boolean
Dim MODIFYFLAG As Boolean
Dim No_Chargeflag As Boolean
Dim NoMemo As Boolean
Dim SELMA10121_QR As rdoQuery
Dim SELMA10121_RS As rdoResultset
Dim SELMA10121FLAG As Boolean
Dim selma104_qr As rdoQuery
Dim selma104_rs As rdoResultset
Dim selma104flag As Boolean
Dim SELTA1011_QR As rdoQuery
Dim SELTA1011_RS As rdoResultset
Dim SELTA1011FLAG As Boolean
Dim SELCLINHIST_QR As rdoQuery
Dim SELCLINHIST_RS As rdoResultset
Dim SELCLINHISTFLAG As Boolean
Dim SELRMODALITY_QR As rdoQuery
Dim SELRMODALITY_RS As rdoResultset
Dim SELRMODALITYFLAG As Boolean
Dim SELRINVEST_QR As rdoQuery
Dim SELRINVEST_RS As rdoResultset
Dim SELRINVESTFLAG As Boolean
Dim SELMA101_QR As rdoQuery
Dim SELMA101_RS As rdoResultset
'Dim SELMA101_FLAG As Boolean
Dim SELMA101FLAG As Boolean
Dim SELMA301_QR As rdoQuery
Dim SELMA301_RS As rdoResultset
Dim SELMA301FLAG As Boolean
Dim SELMA10122_QR As rdoQuery
Dim SELMA10122_RS As rdoResultset
Dim SELMA10122FLAG As Boolean
Dim SELSTAFFINFO_QR As rdoQuery
Dim SELSTAFFINFO_RS As rdoResultset
Dim SELSTAFFINFOFLAG As Boolean

Dim SELEMPBASINFO_QR As rdoQuery
Dim SELEMPBASINFO_RS As rdoResultset
Dim SELEMPBASINFOFLAG As Boolean
'ADDED FOR RTIMELOG
Dim INSTIME_QR As rdoQuery
Dim INSTIME_RS As rdoResultset
Dim INSRREQUISITN_QR As rdoQuery
Dim SELDAUTOREQNO_QR As rdoQuery
Dim SELDAUTOREQNO_RS As rdoResultset
Dim SELDAUTOREQNOFLAG As Boolean
Dim INSDAUTOREQNO_QR As rdoQuery
Dim UPDDAUTOREQNO_QR As rdoQuery
Dim SELDAUTOPATH_QR As rdoQuery
Dim SELDAUTOPATH_RS As rdoResultset
Dim INSDAUTOPATH_QR As rdoQuery
Dim UPDDAUTOPATH_QR As rdoQuery
Dim SELDAUTOPATHFLAG As Boolean
Dim SELRREQUISITN_QR As rdoQuery
Dim SELRREQUISITN_RS As rdoResultset
Dim SELRREQUISITNFLAG As Boolean
Dim INSDREQUISITN_QR As rdoQuery
Dim UPDTA1012_QR As rdoQuery
Dim insclinhist_qr As rdoQuery
Dim updclinhist_qr As rdoQuery
Dim selta1012_qr As rdoQuery
Dim selta1012_rs As rdoResultset
Dim SELTA1012FLAG As Boolean
Dim insta1012_qr As rdoQuery
Dim SELRREQUISITN_QR2 As rdoQuery
Dim SELRREQUISITN_RS2 As rdoResultset
Dim SELRREQUISITNFLAG2 As Boolean
Dim seldrequisitn_qr As rdoQuery
Dim seldrequisitn_rs As rdoResultset
Dim SELDREQUISITNFLAG As Boolean
Dim SELRMODALITY_QR2 As rdoQuery
Dim SELRMODALITY_RS2 As rdoResultset
Dim SELRMODALTIYFLAG2 As Boolean
Dim SELMA106_QR2 As rdoQuery
Dim SELMA106_RS2 As rdoResultset
Dim SELMA106FLAG2 As Boolean
Dim INSIHREQ_QR As rdoQuery
Dim SELTA113_QR As rdoQuery
Dim SELTA113_RS As rdoResultset
Dim SELTA113FLAG As Boolean
Dim SELTA114_QR As rdoQuery
Dim SELTA114_RS As rdoResultset
Dim SELTA114FLAG As Boolean
Dim DPATHYEAR_QR As rdoQuery
Dim DPATHYEAR_RS As rdoResultset
Dim DPATHYEARFLAG As Boolean
Dim INSDMULPROCE_QR As rdoQuery
Dim REPEATFLAG As Boolean
'for crystal code
    Dim AMQR As rdoQuery
    Dim AMRS As rdoResultset
    Dim AMFLAG As Boolean
Dim SELRREQ_QR As rdoQuery
Dim SELRREQ_RS As rdoResultset
Dim SELRREQFLAG As Boolean

Dim SELMA301QR As rdoQuery
Dim SELMA301RS As rdoResultset
Dim SELMA1031_QR As rdoQuery
Dim SELMA1031_RS As rdoResultset
Dim SELMA1031FLAG As Boolean

Dim SELEMPDPDLIST_QR As rdoQuery
Dim SELEMPDPDLIST_RS As rdoResultset
Dim SELEMPDPDLISTFLAG As Boolean

Dim INSDREQSTAFFD_QR As rdoQuery

Dim SELEMPDPDLIST1_QR As rdoQuery
Dim SELEMPDPDLIST1_RS As rdoResultset
Dim SELEMPDPDLIST1FLAG As Boolean
Dim INSDIHREPORT_QR As rdoQuery
Dim seldihreport_qr As rdoQuery
Dim seldihreport_rs As rdoResultset
Dim seldihreportflag As Boolean

Dim SELTA101_QR As rdoQuery
Dim SELTA101_RS As rdoResultset
Dim SELTA101FALG As Boolean
Dim selMA101QR As rdoQuery
Dim SELMA101RS As rdoResultset
Dim SELMA101FLAG1 As Boolean

Dim sitedupqr As rdoQuery
Dim siteduprs As rdoResultset
Dim sitedupflag As Boolean
Dim sitedupqr1 As rdoQuery
Dim siteduprs1 As rdoResultset
Dim siteduprs1flag As Boolean
Dim strins As String
Dim siteinsertqr As rdoQuery
Dim WMETERIAL As Application
Dim DMETERIAL As Object
Dim WMETERIALFN As Application
Dim DMETERIALFN As Object
Dim WMETERIALCN As Application
Dim DMETERIALCN As Object
Dim WMETERIALIH As Application
Dim DMETERIALIH As Object

Dim WSCOPYCY As Application
Dim DSCOPYCY As Object
Dim WTREATCY As Application
Dim DTREATCY As Object
Dim WOUTSPCY As Application
Dim DOUTSPCY As Object

Dim str12 As String
Dim SELDREQQR As rdoQuery
Dim SELDREQRS As rdoResultset
Dim seLdreqrsflag As Boolean
Dim upddrequisitn_qr As rdoQuery
Dim updrrequisitn_qr As rdoQuery
Dim UPDDREQIHQR As rdoQuery
Dim insdsftdcyqr As rdoQuery
Dim upddsftdcyqr As rdoQuery
Dim seldsftdcyqr As rdoQuery
Dim seldsftdcyrs As rdoResultset
Dim seldsftdcyflag As Boolean
Dim SELDENTERQR As rdoQuery
Dim SELDENTERRS As rdoResultset
Dim SELDENTERFLAG As Boolean
Dim UPDRREQQR As rdoQuery
Dim INSSPONANTIBQR As rdoQuery
Dim UPDSPONANTIBQR As rdoQuery
Dim SELSPONANTIBQR As rdoQuery
Dim SELSPONANTIBRS As rdoResultset
Dim SELSPONANTIBFLAG As Boolean
Dim SELDMULPROCEQR As rdoQuery
Dim SELDMULPROCERS As rdoResultset
Dim SELDMULPROCEFLAG As Boolean

Dim PRE_QUERY As rdoQuery
Dim PRE_SERCODE As rdoQuery

Dim TSSAMPLE As String
Dim INSINDICATIONQR As rdoQuery
Dim UPDINDICATIONQR As rdoQuery
Dim SELINDICATIONQR As rdoQuery
Dim SELINDICATIONRS As rdoResultset
Dim SELINDICATIONFLAG As Boolean

Dim INSDBMAINDQR As rdoQuery
Dim SELDBMAINDQR As rdoQuery
Dim SELDBMAINDRS As rdoResultset
Dim SELDBMAINDFLAG As Boolean

Dim INSBLOODPRODUQR As rdoQuery
Dim SELBLOODPRODUQR As rdoQuery
Dim SELBLOODPRODURS As rdoResultset
Dim SELBLOODPRODUFLAG As Boolean

Dim SELDREQNOQR As rdoQuery
Dim SELDREQNORS As rdoResultset
Dim SELDREQNOFLAG As Boolean

Dim seldrequisitn_FLAG As Boolean

''BP
Dim SELDTMNO_QR As rdoQuery
Dim SELDTMNO_RS As rdoResultset
Dim SELDTMNO_FLAG As Boolean

Dim INSDTMNO_QR As rdoQuery

Dim NOCHARG_QR As rdoQuery
Dim NOCHARG_RS As rdoResultset
Dim NOCHARG_FLG As Boolean
''ON 30/05/03

Dim SELRINVEST1_QR As rdoQuery
Dim SELRINVEST1_RS As rdoResultset
Dim SELRINVEST1_FLAG As Boolean

Dim SELMA101_QR1 As rdoQuery
Dim SELMA101_RS1 As rdoResultset
Dim SELMA101_FLAG1 As Boolean

''ON 02/06/03

Dim SELRINVEST11_QR As rdoQuery
Dim SELRINVEST11_RS As rdoResultset
Dim SELRINVEST11_FLAG As Boolean

Dim SELMA101_QR11 As rdoQuery
Dim SELMA101_RS11 As rdoResultset
Dim SELMA101_FLAG11 As Boolean

Dim SELMA101hl_QR As rdoQuery
Dim SELMA101hl_RS As rdoResultset
Dim SELMA101hl_FLAG As Boolean

Dim updreq_qr As rdoQuery

Dim s_charges As String
''emp name
Dim SELRISSTAT_QR As rdoQuery
Dim SELRISSTAT_RS As rdoResultset
Dim SELRISSTAT_FLAG As Boolean

'Dim v_rep As Boolean

'Dim SELDTMTAB_QR As rdoQuery
'Dim SELDTMTAB_RS As rdoResultset
'Dim SELDTMTAB_FLAG As Boolean

''****ICU*******

Dim UPDICU_QR As rdoQuery

Dim SELICU_QR As rdoQuery
Dim SELICU_RS As rdoResultset
Dim SELICU_FLAG As Boolean

'SELDISICU_QR

Dim SELDISICU_QR As rdoQuery
Dim SELDISICU_RS As rdoResultset
Dim SELDISICU_FLAG As Boolean
Dim ICU As String
Dim MCICU As String
Dim MCRWD As String
''****ICU*******

'''**********************
''FOR SP REPORT

Dim seldrequisitn_qr10 As rdoQuery
Dim seldrequisitn_rs10 As rdoResultset
Dim seldrequisitn_flag10 As Boolean

Dim seldrequisitn_qr11 As rdoQuery
Dim seldrequisitn_rs11 As rdoResultset
Dim seldrequisitn_flag11 As Boolean


Dim X
Dim DEPT1 As String
Dim MODALITY1 As String
Dim REQNO1 As String

'********** Shraddha Pandit 21/05/2009 ***********
Dim SEL_MEMOENABL_QR As rdoQuery
Dim SEL_MEMOENABL_RS As rdoResultset
Dim MD As String
'*************************************************

Dim SELRREQNOQR As rdoQuery
Dim SELRREQNORS As rdoResultset
Dim SELRREQNOFLAG As Boolean

Dim SELDADDENDUM_QR As rdoQuery
Dim SELDADDENDUM_RS As rdoResultset
Dim SELDADDENDUM_FLAG As Boolean

Dim SELPATHDREQUISITN_QR As rdoQuery
Dim selpathdrequisitn_rs As rdoResultset

Dim SELDFREPORT_QR As rdoQuery
Dim SELDFREPORT_RS As rdoResultset

'''****************************
''****HL***********************
Dim SELDBMACIND_QR As rdoQuery
Dim SELDBMACIND_RS As rdoResultset
Dim SELDBMACIND_FLAG As Boolean

Dim PREPATH_QR As rdoQuery
Dim PREPATH_RS As rdoResultset
Dim PREPATH_FLAG As Boolean

Dim INSDHEMCYTO_QR As rdoQuery
Dim SELDHEMCYTO_QR1 As rdoQuery
Dim SELDHEMCYTO_RS As rdoResultset
Dim SELDHEMCYTO_FLAG As Boolean

''********HL******************
''*****BP CROSSMATCH*********

Dim UPDREQNSTAT_QR As rdoQuery
Dim UPDDMULPROCE_QR As rdoQuery

'************FOR BP(RDP,SDP,FFP,FVIIID,CRYO,ALB) UPDATE MEMO NO AS NOCHARGE **************
Dim UPDMEMORREQ_QR As rdoQuery
'************FOR BP(RDP,SDP,FFP,FVIIID,CRYO,ALB) UPDATE MEMO NO AS NOCHARGE **************

Dim Dept_cd1 As String
Dim sel_DptCd_qr As rdoQuery
Dim sel_DptCd_rs As rdoResultset
Dim enter_by1 As String
'Dim NM_FLG As Boolean
''*****BP CROSSMATCH*********
'*** inserting into TBPRODUCT of Pimpol
    Dim Ins_TBProduct_qr As rdoQuery
'*** inserting into TBPRODUCT of Pimpol


Dim g253cd, g106cd, g119cd As Integer

'''ADDED BY KANCHAN FOR BARCODE
Dim SELDEPTDES_QR As rdoQuery
Dim SELDEPTDES_RS As rdoResultset
Dim VIEWFLAG_BCD As Boolean
Dim NEWFLAG_BCD As Boolean
Dim colorcd As String
Dim smplcd As String
Dim DESCRIP As String
Dim UPDRTIMEFORBC As rdoQuery

'' ADDED BY SANDHYA 1.8.2008 FOR PROTOCOL PATIENT
Dim SELprcase_QR As rdoQuery
Dim SELPRCASE_RS As rdoResultset
Dim PROTSERV_QR As rdoQuery
Dim PROTSERV_RS As rdoResultset
Dim QR_PRCASE As rdoQuery
Dim RS_PRCASE As rdoResultset
Dim prflg As Boolean





Sub SELMA10121()
''STR = "SELECT * FROM PATOOLSLIB.MA10121"
''Set SELMA10121_QR = rdb.CreateQuery("SELMA1012111_QR11", STR)
End Sub



Private Sub cbo_proccode_Click()
CMB_CGPROCEDURE.ListIndex = cbo_proccode.ListIndex
End Sub



Private Sub CMB_CGNIP_GotFocus()
CMB_CGNIP.clear
CMB_CGNIP.AddItem "VPP"
CMB_CGNIP.AddItem "LR"
CMB_CGNIP.AddItem "BAR"
CMB_CGNIP.AddItem "DrOp"
End Sub

Private Sub CMB_CGPROCEDURE_Click()
cbo_proccode.ListIndex = CMB_CGPROCEDURE.ListIndex
End Sub

Private Sub cmb_cpreinv_Click()
  X = InStr(1, Trim(cmb_cpreinv.Text), "$")
  REQNO1 = Mid(Trim(cmb_cpreinv.Text), 1, X - 1)
  DEPT1 = Mid(Trim(REQNO1), 1, 3)
  MODALITY1 = Mid(Trim(REQNO1), 4, 2)
  If DEPT1 = "FZZ" And MODALITY1 = "SP" Then
    SPREPORT
  ElseIf DEPT1 = "FZZ" And MODALITY1 = "FA" Then
     FAREPORT
  ElseIf DEPT1 = "FZZ" And (MODALITY1 = "CG" Or MODALITY1 = "CN") Then
     CGCNREPORT
  ElseIf DEPT1 = "LZZ" And (MODALITY1 = "MM" Or MODALITY1 = "MR" Or MODALITY1 = "CR" Or MODALITY1 = "CT" Or MODALITY1 = "US") Then
     RISREPORT
 End If
End Sub


Private Sub cmb_cyoutspecimen_GotFocus()
On Error GoTo m1:
WOUTSPCY.Visible = False
Set DOUTSPCY = WOUTSPCY.ActiveDocument
DOUTSPCY.Content = ""
DOUTSPCY.Content = cmb_cyoutspecimen
WOUTSPCY.Caption = "Outside Specimen"
WOUTSPCY.WindowState = wdWindowStateMaximize
DOUTSPCY.Content = cmb_cyoutspecimen
WOUTSPCY.Visible = True
WOUTSPCY.Activate
m1:
On Error Resume Next
Set WOUTSPCY = CreateObject("Word.Application")
WOUTSPCY.Documents.add ("C:\DISREPORTSEC\REPORT.DOT")
WOUTSPCY.Caption = "Outside Specimen"
WOUTSPCY.WindowState = wdWindowStateMaximize
Set DOUTSPCY = WOUTSPCY.ActiveDocument
DOUTSPCY.Content = Trim(cmb_cyoutspecimen.Text)
WOUTSPCY.Visible = True
WOUTSPCY.Activate
End Sub

Private Sub cmb_cyoutspecimen_LostFocus()
On Error GoTo m2:
Set DOUTSPCY = WOUTSPCY.ActiveDocument
cmb_cyoutspecimen.Text = DOUTSPCY.Content
If Len(Trim(cmb_cyoutspecimen.Text)) > 500 Then
    MsgBox "Out side specimen Max of 500 Characters"
    cmb_cyoutspecimen.SetFocus
    Exit Sub
End If
'flagimp = False
DOUTSPCY.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'DOUTSPCY.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
DOUTSPCY.Close (wdDoNotSaveChanges)
WOUTSPCY.quit
'W1.Visible = False
Exit Sub
m2:
MsgBox Err.Description
End Sub

Private Sub cmb_cyprocedure_Click()
cmb_cyprocode.ListIndex = cmb_cyprocedure.ListIndex
End Sub

Private Sub cmb_cyprocode_Click()
cmb_cyprocedure.ListIndex = cmb_cyprocode.ListIndex
End Sub

Private Sub cmb_dependents_Click()
Dim K As Integer
On Error GoTo MSG:
If SELEMPDPDLIST1FLAG = False Then
    SELEMPDPDLIST1_QR(0) = Trim(Msk_caseno.Text)
    I = InStr(cmb_dependents, "$")
    K = Len(cmb_dependents.Text)
    SELEMPDPDLIST1_QR(1) = Trim(Mid(cmb_dependents, I + 1, K - I))
    Set SELEMPDPDLIST1_RS = SELEMPDPDLIST1_QR.OpenResultset(1, 3)
    SELEMPDPDLIST1FLAG = True
Else
    SELEMPDPDLIST1_QR(0) = Trim(Msk_caseno.Text)
    I = InStr(cmb_dependents, "$")
    SELEMPDPDLIST1_QR(1) = Trim(Mid(cmb_dependents, I + 1, 10))
    SELEMPDPDLIST1_RS.MoreResults
    SELEMPDPDLIST1_RS.Requery
End If

'ADDED BY MILIND ON 02-07-2006
While SELEMPDPDLIST1_RS.EOF = False
    Dim n
    n = Mid(cmb_dependents, 1, I - 1)
    If Trim(SELEMPDPDLIST1_RS("DPDNAME")) = n Then
    
        txt_name = Left(cmb_dependents, I - 1)
        txt_sex = SELEMPDPDLIST1_RS("dpdsex")
        txt_Age = DateDiff("yyyy", SELEMPDPDLIST1_RS("DPDDOB"), DATE11)
    
    End If
        SELEMPDPDLIST1_RS.MoveNext
Wend
'UPTO HERE

Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub cmb_dept_Click()

Dim I As Integer
On Error GoTo MSG:
I = InStr(cmb_dept.Text, "$")
txt_dept.Text = Mid(cmb_dept.Text, I + 1, 4)
'If v_rep = True And Trim(txt_dept.Text) = "FEC" Then
'    MsgBox "Please Don't Repeat the Requisitn !!", vbInformation, "Requisition"
'    Exit Sub
'End If

If SELRMODALITYFLAG = False Then
    SELRMODALITY_QR(0) = Trim(txt_dept.Text)
    Set SELRMODALITY_RS = SELRMODALITY_QR.OpenResultset(1, 3)
    SELRMODALITYFLAG = True
 Else
    SELRMODALITY_QR(0) = Trim(txt_dept.Text)
    SELRMODALITY_RS.MoreResults
    SELRMODALITY_RS.Requery
End If
If SELRMODALITY_RS.RowCount > 0 Then
    cmb_modality.clear
    SELRMODALITY_RS.MoveFirst
  
    While Not SELRMODALITY_RS.EOF
    If Trim(Trim(SELRMODALITY_RS("MODALITYCD"))) <> "IH" Then
    cmb_modality.AddItem Trim(SELRMODALITY_RS("MODLITYDES")) & "$" & Trim(SELRMODALITY_RS("MODALITYCD"))
    End If
    If Trim(txt_dept.Text) = "CBZ" Then
    cmb_modality.Text = Trim(SELRMODALITY_RS("MODLITYDES")) & "$" & Trim(SELRMODALITY_RS("MODALITYCD"))
    txt_modcode.Text = Trim(SELRMODALITY_RS("MODALITYCD"))
    End If
    SELRMODALITY_RS.MoveNext
  
    Wend
End If


Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub cmb_dept_GotFocus()
On Error GoTo MSG:
Dim ind_cbodept As Integer
If Trim(cmb_deptname) = "" Then
    For ind_cbodept = 0 To (cmb_dept.ListCount - 1)
        If Trim(cmb_dept.List(ind_cbodept)) = "PATHOLOGY$FZZ" Then
            cmb_dept.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub



Private Sub cmb_edprereqno_Click()
'''X = InStr(1, Trim(cmb_prereqno.Text), "$")
'''REQNO1 = Mid(Trim(cmb_prereqno.Text), 1, X - 1)
'''DEPT1 = Mid(Trim(REQNO1), 1, 3)
'''MODALITY1 = Mid(Trim(REQNO1), 4, 2)
'''If DEPT1 = "FZZ" And MODALITY1 = "SP" Then
'''    SPREPORT
'''ElseIf DEPT1 = "FZZ" And MODALITY1 = "FA" Then
'''    FAREPORT
'''ElseIf DEPT1 = "FZZ" And (MODALITY1 = "CG" Or MODALITY1 = "CN") Then
'''    CGCNREPORT
'''ElseIf DEPT1 = "LZZ" And (MODALITY1 = "MM" Or MODALITY1 = "MR" Or MODALITY1 = "CR" Or MODALITY1 = "CT" Or MODALITY1 = "US") Then
'''    RISREPORT
'''End If
End Sub

Private Sub CMB_FANIP_GotFocus()
CMB_FANIP.clear
CMB_FANIP.AddItem "VPP"
CMB_FANIP.AddItem "LR"
CMB_FANIP.AddItem "BAR"
CMB_FANIP.AddItem "DrOp"
End Sub


Private Sub cmb_fnpreinvest_Click()
If Trim(cmb_fnpreinvest.Text) = "" Then
    Exit Sub
End If
    X = InStr(1, Trim(cmb_fnpreinvest.Text), "$")
REQNO1 = Mid(Trim(cmb_fnpreinvest.Text), 1, X - 1)
DEPT1 = Mid(Trim(REQNO1), 1, 3)
MODALITY1 = Mid(Trim(REQNO1), 4, 2)
If DEPT1 = "FZZ" And MODALITY1 = "SP" Then
    SPREPORT
ElseIf DEPT1 = "FZZ" And MODALITY1 = "FA" Then
    FAREPORT
ElseIf DEPT1 = "FZZ" And (MODALITY1 = "CG" Or MODALITY1 = "CN") Then
    CGCNREPORT
ElseIf DEPT1 = "LZZ" And (MODALITY1 = "MM" Or MODALITY1 = "MR" Or MODALITY1 = "CR" Or MODALITY1 = "CT" Or MODALITY1 = "US") Then
    RISREPORT
End If
End Sub

Private Sub CMB_FNPROCEDURE_Click()
cmb_fnprocode.ListIndex = CMB_FNPROCEDURE.ListIndex
End Sub



Private Sub cmb_fnprocode_Click()
CMB_FNPROCEDURE.ListIndex = cmb_fnprocode.ListIndex
End Sub

Private Sub cmb_fnsite_KeyPress(KeyAscii As Integer)
If cmb_fnsite.Text <> "" Then
    KeyAscii = 0
End If
End Sub

Private Sub cmb_frozensec_CLICK()
If cmb_frozensec.Text = "YES" Then
MSK_FROSECDT.Enabled = True
MSK_FROSECDT.Text = Format(DATE11, "dd/mm/yyyy")
Else
MSK_FROSECDT.Enabled = False
End If
End Sub

Private Sub cmb_frozensec_GotFocus()
On Error GoTo MSG:
Dim ind_cbodept As Integer
If Trim(cmb_frozensec) = "" Then
    For ind_cbodept = 0 To (cmb_frozensec.ListCount - 1)
        If Trim(cmb_frozensec.List(ind_cbodept)) = "NO" Then
            cmb_frozensec.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Private Sub cmb_hpprocedure_Click()
  cmb_hpprocedurecode.ListIndex = cmb_hpprocedure.ListIndex
End Sub

Private Sub cmb_hpprocedure_GotFocus()
cmb_hpprocedure.Width = 4500
End Sub

Private Sub cmb_hpprocedure_LostFocus()
cmb_hpprocedure.Width = 3255
End Sub

Private Sub cmb_hpprocedurecode_Click()
cmb_hpprocedure.ListIndex = cmb_hpprocedurecode.ListIndex
End Sub

Private Sub CMB_IHNIP_GotFocus()
CMB_IHNIP.clear
CMB_IHNIP.AddItem "VPP"
CMB_IHNIP.AddItem "LR"
CMB_IHNIP.AddItem "BAR"
CMB_IHNIP.AddItem "DrOp"
End Sub

Private Sub cmb_ihprocedure_Click()
cmb_ihprocode.ListIndex = cmb_ihprocedure.ListIndex
End Sub

Private Sub cmb_ihprocode_Click()
cmb_ihprocedure.ListIndex = cmb_ihprocode.ListIndex
End Sub

Private Sub cmb_Ihsoutsideblock_GotFocus()
On Error GoTo MSG:
Dim ind_cbodept As Integer
If Trim(cmb_Ihsoutsideblock) = "" Then
    For ind_cbodept = 0 To (cmb_Ihsoutsideblock.ListCount - 1)
        If Trim(cmb_Ihsoutsideblock.List(ind_cbodept)) = 0 Then
            cmb_Ihsoutsideblock.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Private Sub cmb_Ihsoutslide_GotFocus()
On Error GoTo MSG:
Dim ind_cbodept As Integer
If Trim(cmb_Ihsoutslide) = "" Then
    For ind_cbodept = 0 To (cmb_Ihsoutslide.ListCount - 1)
        If Trim(cmb_Ihsoutslide.List(ind_cbodept)) = 0 Then
            cmb_Ihsoutslide.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Private Sub cmb_Ihspreinv_Click()
''    X = InStr(1, Trim(cmb_prereqno.Text), "$")
''REQNO1 = Mid(Trim(cmb_prereqno.Text), 1, X - 1)
''DEPT1 = Mid(Trim(REQNO1), 1, 3)
''MODALITY1 = Mid(Trim(REQNO1), 4, 2)
''If DEPT1 = "FZZ" And MODALITY1 = "SP" Then
''    SPREPORT
''ElseIf DEPT1 = "FZZ" And MODALITY1 = "FA" Then
''    FAREPORT
''ElseIf DEPT1 = "FZZ" And (MODALITY1 = "CG" Or MODALITY1 = "CN") Then
''    CGCNREPORT
''ElseIf DEPT1 = "LZZ" And (MODALITY1 = "MM" Or MODALITY1 = "MR" Or MODALITY1 = "CR" Or MODALITY1 = "CT" Or MODALITY1 = "US") Then
''    RISREPORT
''End If
End Sub

Private Sub cmb_mbprevunv_Click()
''    X = InStr(1, Trim(cmb_prereqno.Text), "$")
''REQNO1 = Mid(Trim(cmb_prereqno.Text), 1, X - 1)
''DEPT1 = Mid(Trim(REQNO1), 1, 3)
''MODALITY1 = Mid(Trim(REQNO1), 4, 2)
''If DEPT1 = "FZZ" And MODALITY1 = "SP" Then
''    SPREPORT
''ElseIf DEPT1 = "FZZ" And MODALITY1 = "FA" Then
''    FAREPORT
''ElseIf DEPT1 = "FZZ" And (MODALITY1 = "CG" Or MODALITY1 = "CN") Then
''    CGCNREPORT
''ElseIf DEPT1 = "LZZ" And (MODALITY1 = "MM" Or MODALITY1 = "MR" Or MODALITY1 = "CR" Or MODALITY1 = "CT" Or MODALITY1 = "US") Then
''    RISREPORT
''End If
End Sub

Private Sub CMB_MEMO_Click()
Dim I As Integer
On Error GoTo MSG:

MEMOFLAG = True
'ASSREQNO (AA)
If Trim(txt_status) = "IP" And INNC = False Then
        
        ip_memo_d1.me_date.Text = Format(DATE11, "dd/mm/yyyy")
        I = InStr(cmb_dept.Text, "$")
        ip_memo_d1.Com_dno.AddItem (cmb_dept.Text)
        ip_memo_d1.Com_dno.Text = (cmb_dept.Text)
        'op_memo_d1.Com_dno.AddItem Left(cmb_dept.Text, I - 1)
        'op_memo_d1.Com_dno.Text = Left(cmb_dept.Text, I - 1)
        ip_memo_d1.lbl_wf_bed = txt_ward
        
        ip_memo_d1.lbl_dno = txt_dept.Text
        ip_memo_d1.me_cno.Text = Msk_caseno.FormattedText
        ip_memo_d1.lbl_name = txt_name.Text
        ip_memo_d1.lbl_cat = txt_catgory.Text
        ip_memo_d1.lbl_sex = txt_sex
        ip_memo_d1.lbl_age = txt_Age.Text
        ip_memo_d1.lbl_typ = PATYPE
        ip_memo_d1.txt_unit_doc_name = txt_pridocser.Text
        ip_memo_d1.me_rdate.Text = Format(DATE11, "DD/MM/YYYY")
        ip_memo_d1.commit.Visible = True
        ip_memo_d1.cancel.Visible = True
        ip_memo_d1.print1.Visible = False
        ip_memo_d1.d_u_name.Enabled = True
        '' Sandhya 1.8.2008 to search protocol patient
                        
          SELprcase_QR(0) = Msk_caseno.FormattedText
          Set SELPRCASE_RS = SELprcase_QR.OpenResultset(1, 3)
          If SELPRCASE_RS.RowCount > 0 Then
             ip_memo_d1.Label7.Enabled = True
             ip_memo_d1.Label7.Visible = True
             ip_memo_d1.TXTNET.Enabled = True
             ip_memo_d1.TXTNET.Visible = True
             ip_memo_d1.TXTPRTTOT.Enabled = True
             ip_memo_d1.TXTPRTTOT.Visible = True
             ip_memo_d1.Label19.Enabled = True
             ip_memo_d1.Label19.Visible = True
             ip_memo_d1.Label20.Enabled = True
             ip_memo_d1.Label20.Visible = True
              MIRBNO = CMBIRB.Text
          
             SELPRCASEFLG = True
          Else
             SELPRCASEFLG = False
             ip_memo_d1.Label7.Enabled = False
             ip_memo_d1.Label7.Visible = False
             ip_memo_d1.TXTNET.Enabled = False
             ip_memo_d1.TXTNET.Visible = False
             ip_memo_d1.TXTPRTTOT.Enabled = False
             ip_memo_d1.TXTPRTTOT.Visible = False
             ip_memo_d1.Label19.Enabled = False
             ip_memo_d1.Label19.Visible = False
             ip_memo_d1.Label20.Enabled = False
             ip_memo_d1.Label20.Visible = False
             MIRBNO = ""
          End If
          
          '' end of protocol patient
        If SELTA113FLAG = False Then
            SELTA113_QR(0) = Msk_caseno.FormattedText
            Set SELTA113_RS = SELTA113_QR.OpenResultset(1, 3)
            SELTA113FLAG = True
        Else
            
            SELTA113_QR(0) = Msk_caseno.FormattedText
            SELTA113_RS.MoreResults
            SELTA113_RS.Requery
        End If
        If SELTA113_RS.RowCount > 0 Then
            ip_memo_d1.lbl_tr_no = SELTA113_RS("tr_acc_no")
        End If
        If selta114falg = False Then
            SELTA114_QR(0) = Msk_caseno.FormattedText
            Set SELTA114_RS = SELTA114_QR.OpenResultset(1, 3)
            SELTA114FLAG = True
        Else
            SELTA114_QR(0) = Msk_caseno.FormattedText
            SELTA114_RS.MoreResults
            SELTA114_RS.Requery
        End If
        If SELTA114_RS.RowCount > 0 Then
            ip_memo_d1.lbl_cr_card_no = SELTA114_RS("cr_card")
        End If
            
        


ip_memo_d1.show
IPOPMEMO = True
End If
If txt_status = "OP" Then

        op_memo_d1.me_date.Text = Format(DATE11, "dd/mm/yyyy")
        I = InStr(cmb_dept.Text, "$")
        op_memo_d1.Com_dno.AddItem (cmb_dept.Text)
        op_memo_d1.Com_dno.Text = (cmb_dept.Text)
        'op_memo_d1.Com_dno.AddItem Left(cmb_dept.Text, I - 1)
        'op_memo_d1.Com_dno.Text = Left(cmb_dept.Text, I - 1)
        op_memo_d1.lbl_dname = txt_dept.Text
        op_memo_d1.me_cno.Text = Msk_caseno.FormattedText
        op_memo_d1.lbl_name = txt_name.Text
        op_memo_d1.lbl_cat = txt_catgory.Text
        op_memo_d1.lbl_sex = txt_sex
        op_memo_d1.lbl_age = txt_Age.Text
        op_memo_d1.lbl_typ = PATYPE
        op_memo_d1.txt_unit_doc_name = txt_pridocser.Text
        op_memo_d1.me_rdate.Text = Format(DATE11, "DD/MM/YYYY")
        op_memo_d1.commit.Visible = True
        op_memo_d1.cancel.Visible = True
        op_memo_d1.print1.Visible = False
        op_memo_d1.d_u_name.Enabled = True
        '' Added by Sandhya 1.8.2008 to search protocol patient
        SELprcase_QR(0) = Msk_caseno.FormattedText
          Set SELPRCASE_RS = SELprcase_QR.OpenResultset(1, 3)
          If SELPRCASE_RS.RowCount > 0 Then
          '   op_memo_d1.lbl_tr_no = SELPRCASE_RS("PCODE")
             SELPRCASEFLG = True
             op_memo_d1.Label15.Enabled = True
             op_memo_d1.Label15.Visible = True
             op_memo_d1.nettot.Enabled = True
             op_memo_d1.nettot.Visible = True
             op_memo_d1.prottot.Enabled = True
             op_memo_d1.prottot.Visible = True
             op_memo_d1.Label6.Enabled = True
             op_memo_d1.Label6.Visible = True
             op_memo_d1.Label17.Enabled = True
             op_memo_d1.Label17.Visible = True
             MIRBNO = CMBIRB.Text
          Else
             SELPRCASEFLG = False
             op_memo_d1.Label15.Enabled = False
             op_memo_d1.Label15.Visible = False
             op_memo_d1.nettot.Enabled = False
             op_memo_d1.nettot.Visible = False
             op_memo_d1.prottot.Enabled = False
             op_memo_d1.prottot.Visible = False
             op_memo_d1.Label6.Enabled = False
             op_memo_d1.Label6.Visible = False
             op_memo_d1.Label17.Enabled = False
             op_memo_d1.Label17.Visible = False
             MIRBNO = ""
          End If
          '' end of protcol patient
        If SELTA113FLAG = False Then
            SELTA113_QR(0) = Msk_caseno.FormattedText
            Set SELTA113_RS = SELTA113_QR.OpenResultset(1, 3)
            SELTA113FLAG = True
        Else
            
            SELTA113_QR(0) = Msk_caseno.FormattedText
            SELTA113_RS.MoreResults
            SELTA113_RS.Requery
        End If
        If SELTA113_RS.RowCount > 0 Then
            op_memo_d1.lbl_trust_no = SELTA113_RS("tr_acc_no")
        End If
        If selta114falg = False Then
            SELTA114_QR(0) = Msk_caseno.FormattedText
            Set SELTA114_RS = SELTA114_QR.OpenResultset(1, 3)
            SELTA114FLAG = True
        Else
            SELTA114_QR(0) = Msk_caseno.FormattedText
            SELTA114_RS.MoreResults
            SELTA114_RS.Requery
        End If
        If SELTA114_RS.RowCount > 0 Then
            op_memo_d1.lbl_cr_card_no = SELTA114_RS("cr_card")
        End If
        
         
        
        IPOPMEMO = True
        op_memo_d1.show
End If
If txt_status = "IP" And INNC = True Then
    op_memo_d1.me_date.Text = Format(DATE11, "dd/mm/yyyy")
        I = InStr(cmb_dept.Text, "$")
        op_memo_d1.Com_dno.AddItem (cmb_dept.Text)
        op_memo_d1.Com_dno.Text = (cmb_dept.Text)
        'op_memo_d1.Com_dno.AddItem Left(cmb_dept.Text, I - 1)
        'op_memo_d1.Com_dno.Text = Left(cmb_dept.Text, I - 1)
        op_memo_d1.lbl_dname = txt_dept.Text
        op_memo_d1.me_cno.Text = Msk_caseno.FormattedText
        op_memo_d1.lbl_name = txt_name.Text
        op_memo_d1.lbl_cat = txt_catgory.Text
        op_memo_d1.lbl_sex = txt_sex
        op_memo_d1.lbl_age = txt_Age.Text
        op_memo_d1.lbl_typ = PATYPE
        op_memo_d1.txt_unit_doc_name = txt_pridocser.Text
        op_memo_d1.me_rdate.Text = Format(DATE11, "DD/MM/YYYY")
         '' Added by Sandhya 1.8.2008 to search protocol patient
        SELprcase_QR(0) = Msk_caseno.FormattedText
          Set SELPRCASE_RS = SELprcase_QR.OpenResultset(1, 3)
          If SELPRCASE_RS.RowCount > 0 Then
             SELPRCASEFLG = True
             op_memo_d1.Label15.Enabled = True
             op_memo_d1.Label15.Visible = True
             op_memo_d1.nettot.Enabled = True
             op_memo_d1.nettot.Visible = True
             op_memo_d1.prottot.Enabled = True
             op_memo_d1.prottot.Visible = True
             op_memo_d1.Label6.Enabled = True
             op_memo_d1.Label6.Visible = True
             op_memo_d1.Label17.Enabled = True
             op_memo_d1.Label17.Visible = True
             MIRBNO = CMBIRB.Text
           Else
             SELPRCASEFLG = False
             op_memo_d1.Label15.Enabled = False
             op_memo_d1.Label15.Visible = False
             op_memo_d1.nettot.Enabled = False
             op_memo_d1.nettot.Visible = False
             op_memo_d1.prottot.Enabled = False
             op_memo_d1.prottot.Visible = False
             op_memo_d1.Label6.Enabled = False
             op_memo_d1.Label6.Visible = False
             op_memo_d1.Label17.Enabled = False
             op_memo_d1.Label17.Visible = False
             MIRBNO = ""
          End If
          '' end of protocol
        If SELTA113FLAG = False Then
            SELTA113_QR(0) = Msk_caseno.FormattedText
            Set SELTA113_RS = SELTA113_QR.OpenResultset(1, 3)
            SELTA113FLAG = True
        Else
            
            SELTA113_QR(0) = Msk_caseno.FormattedText
            SELTA113_RS.MoreResults
            SELTA113_RS.Requery
        End If
        If SELTA113_RS.RowCount > 0 Then
            op_memo_d1.lbl_trust_no = SELTA113_RS("tr_acc_no")
        End If
        If selta114falg = False Then
            SELTA114_QR(0) = Msk_caseno.FormattedText
            Set SELTA114_RS = SELTA114_QR.OpenResultset(1, 3)
            SELTA114FLAG = True
        Else
            SELTA114_QR(0) = Msk_caseno.FormattedText
            SELTA114_RS.MoreResults
            SELTA114_RS.Requery
        End If
        If SELTA114_RS.RowCount > 0 Then
            op_memo_d1.lbl_cr_card_no = SELTA114_RS("cr_card")
        End If
        
         
        
        IPOPMEMO = True
        op_memo_d1.show
End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub
Function ASSREQNO(STRNO As String)
SHW_REQNO = Data1.Recordset.Fields("REQNO")
'SHW_REQNO = cmb_reqno.Text
End Function
Private Sub cmb_modality_Click()
Dim I As Integer
On Error GoTo MSG:
I = InStr(cmb_modality.Text, "$")
txt_modcode.Text = Trim(Mid(cmb_modality.Text, I + 1, 2))
If txt_modcode.Text = "SP" Then
   ' NM_FLG = True
    SP
    SPQUERY
End If
''  procedure for cytogentics modality requisition  13-04-2005
 If txt_modcode.Text = "CC" Or txt_modcode.Text = "FI" Then
   cytogeneticreq.clear
   StrCgdiag1 = ""
   StrCgdiag2 = ""
   CY
 End If
 
 If txt_modcode.Text = "MP" Then
   MP
 End If

''  procedure for cytogentics modality requisition 13-04-2005
If txt_modcode.Text = "IH" Then   '**** commented by chandu 3/01/2003
      MsgBox "Please select the Surgical Pathology Modality "
'      btnenabl ("INI")
'      RCLEAR
''    IMHrequisation.Left = pathology.Left + 100
''    IMHrequisation.Top = pathology.Top + 5250
''    IMHrequisation.show
''    IHCLEAR
''    txt_ihpathno.Visible = False
''    lbl_ihpathno.Visible = False
''    If IHFLAG = True Then
''    txt_sppathreqno.Text = sppathreport.cmb_reqno.Text
''    txt_sppathreqno.Enabled = False
''    txt_ihpathno.Visible = True
''    lbl_ihpathno.Visible = True
''    txt_ihpathno.Text = sppathreport.TXT_PATH.Text
''    txt_ihpathno.Enabled = False
''    End If
    'IH
End If
If txt_modcode.Text = "CG" Or txt_modcode.Text = "CN" Or txt_modcode.Text = "PN" Or txt_modcode.Text = "PG" Then
    cg
    SPQUERY
End If

If txt_modcode.Text = "FA" Then
    FA
    SPQUERY
End If
If txt_modcode.Text = "BC" Then
    BC
End If
If txt_modcode.Text = "EL" Then
    EL
End If
If txt_modcode.Text = "TM" Then
    TM
End If
If txt_modcode.Text = "BA" Or txt_modcode.Text = "MB" Or txt_modcode.Text = "SE" Or txt_modcode = "MC" Or txt_modcode.Text = "CM" Or txt_modcode.Text = "MO" Or txt_modcode.Text = "XX" Then
    MB
End If
If txt_modcode.Text = "HL" Then
    HL
End If
If txt_modcode.Text = "BM" Then
    BM
End If
If txt_modcode.Text = "BF" Then
    bfmode = "NEW"
    BF
End If
If txt_modcode.Text = "ED" Then
    ED
End If
If txt_modcode.Text = "2D" Or txt_modcode.Text = "EK" Or txt_modcode.Text = "PF" Or txt_modcode.Text = "ST" Then
    EC
End If
If txt_modcode.Text = "SS" Then
    SS
End If
If txt_modcode.Text = "BP" Then
    BP
End If
If txt_modcode.Text = "HT" Then
    HT
End If
'If cmb_modality.Text = "CY" Then
'surgicalPathology.Visible = False
'immunohistochemistry.Visible = False
'frm_hematopathology.Visible = False
'biochemistry.Visible = False
'cardiology.Visible = False
'Microbiology.Visible = False
'CYTOLOGY.Visible = True
'cytologyfnac.Visible = False
'cpl.Visible = False
'TUMOURMARKER.Visible = False
'HLATYPING.Visible = False
'lbl_req.Left = 3500
'lbl_req.Caption = "REQUISITION - CYTOLOGY"
'pathology.Caption = "Pathology - Cytology"
'End If
'If cmb_modality.Text = "FN" Then
'surgicalPathology.Visible = False
'immunohistochemistry.Visible = False
'frm_hematopathology.Visible = False
'biochemistry.Visible = False
'cardiology.Visible = False
'Microbiology.Visible = False
'CYTOLOGY.Visible = False
'cytologyfnac.Visible = True
'cpl.Visible = False
'TUMOURMARKER.Visible = False
'HLATYPING.Visible = False
'lbl_req.Left = 3500
'lbl_req.Caption = "REQUISITION - FNAC"
'pathology.Caption = "Pathology - FNAC"
'End If
'If cmb_modality.Text = "BC" Then
'surgicalPathology.Visible = False
'immunohistochemistry.Visible = False
'frm_hematopathology.Visible = False
'biochemistry.Show
'cardiology.Visible = False
'Microbiology.Visible = False
'CYTOLOGY.Visible = False
'cytologyfnac.Visible = False
'HLATYPING.Visible = False
'cpl.Visible = False
'lbl_req.Left = 3500
'lbl_req.Caption = "REQUISITION - BIOCHEMISTRY "
'pathology.Caption = " Biochemistry "
'End If
'If cmb_modality.Text = "MB" Then
'surgicalPathology.Visible = False
'immunohistochemistry.Visible = False
'frm_hematopathology.Visible = False
'biochemistry.Visible = False
'cardiology.Visible = False
'Microbiology.Visible = True
'CYTOLOGY.Visible = False
'cytologyfnac.Visible = False
'cpl.Visible = False
'TUMOURMARKER.Visible = False
'HLATYPING.Visible = False
'lbl_req.Left = 3500
'lbl_req.Caption = "REQUISITION - MICROBIOLOGY"
'pathology.Caption = "Pathology - Microbiology"
'End If
'If cmb_modality.Text = "EC" Then
'surgicalPathology.Visible = False
'immunohistochemistry.Visible = False
'frm_hematopathology.Visible = False
'biochemistry.Visible = False
'cardiology.Visible = True
'Microbiology.Visible = False
'CYTOLOGY.Visible = False
'cytologyfnac.Visible = False
'cpl.Visible = False
'TUMOURMARKER.Visible = False
'HLATYPING.Visible = False
'lbl_req.Left = 3500
'lbl_req.Caption = "REQUISITION - CARDIOLOGY "
'pathology.Caption = "Pathology - Cardiology"
'End If
'If cmb_modality.Text = "HP" Then
'surgicalPathology.Visible = False
'immunohistochemistry.Visible = False
'frm_hematopathology.Visible = True
'biochemistry.Visible = False
'cardiology.Visible = False
'Microbiology.Visible = False
'CYTOLOGY.Visible = False
'cytologyfnac.Visible = False
'cpl.Visible = False
'TUMOURMARKER.Visible = False
'HLATYPING.Visible = False
'lbl_req.Left = 3500
'lbl_req.Caption = "REQUISITION - HEMATOPATHOLOGY"
'pathology.Caption = "Pathology - Hematopathology"
'End If
'If cmb_modality.Text = "CL" Then
'surgicalPathology.Visible = False
'immunohistochemistry.Visible = False
'frm_hematopathology.Visible = False
'biochemistry.Visible = False
'cardiology.Visible = False
'Microbiology.Visible = False
'CYTOLOGY.Visible = False
'cytologyfnac.Visible = False
'cpl.Show
'TUMOURMARKER.Visible = False
'HLATYPING.Visible = False
'lbl_req.Left = 3500
'lbl_req.Caption = "REQUISITION - CLINICAL LABORATORY"
'pathology.Caption = "Pathology - Clinical Laboratory"
'End If
'If cmb_modality.Text = "TM" Then
'surgicalPathology.Visible = False
'immunohistochemistry.Visible = False
'frm_hematopathology.Visible = False
'biochemistry.Visible = False
'cardiology.Visible = False
'Microbiology.Visible = False
'CYTOLOGY.Visible = False
'cytologyfnac.Visible = False
'cpl.Visible = False
'TUMOURMARKER.Show
'HLATYPING.Visible = False
'lbl_req.Left = 3500
'lbl_req.Caption = "REQUISITION - TUMOUR MARKER"
'pathology.Caption = "Pathology - Tumour Marker"
'End If
'If cmb_modality.Text = "HL" Then
'surgicalPathology.Visible = False
'immunohistochemistry.Visible = False
'frm_hematopathology.Visible = False
'biochemistry.Visible = False
'cardiology.Visible = False
'Microbiology.Visible = False
'CYTOLOGY.Visible = False
'cytologyfnac.Visible = False
'cpl.Visible = False
'TUMOURMARKER.Visible = False
'HLATYPING.Visible = True
'lbl_req.Left = 3500
'lbl_req.Caption = "REQUISITION - HLA TYPING"
'pathology.Caption = "Pathology - HLA Typing"
'End If

'Procedure.Show
'bloodProducts.txt_dtmno.SetFocus
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub cmb_next_Click()
IMHrequisation.Left = pathology.Left + 100
IMHrequisation.Top = pathology.Top + 5120
IMHrequisation.show

End Sub

Private Sub cmb_modality_LostFocus()
If txt_modcode.Text = "ED" Then
    ED
End If
End Sub

Private Sub cmb_outblock_GotFocus()
On Error GoTo MSG:
Dim ind_cbodept As Integer
If Trim(cmb_outblock) = "" Then
    For ind_cbodept = 0 To (cmb_outblock.ListCount - 1)
        If Trim(cmb_outblock.List(ind_cbodept)) = 0 Then
            cmb_outblock.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Private Sub cmb_outslide_GotFocus()
On Error GoTo MSG:
Dim ind_cbodept As Integer
If Trim(cmb_outslide) = "" Then
    For ind_cbodept = 0 To (cmb_outslide.ListCount - 1)
        If Trim(cmb_outslide.List(ind_cbodept)) = 0 Then
            cmb_outslide.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub


Private Sub CMB_PREINVEST_Click()
X = InStr(1, Trim(cmb_preinvest.Text), "$")
REQNO1 = Mid(Trim(cmb_preinvest.Text), 1, X - 1)
DEPT1 = Mid(Trim(REQNO1), 1, 3)
MODALITY1 = Mid(Trim(REQNO1), 4, 2)
If DEPT1 = "FZZ" And MODALITY1 = "SP" Then
    SPREPORT
ElseIf DEPT1 = "FZZ" And MODALITY1 = "FA" Then
    FAREPORT
ElseIf DEPT1 = "FZZ" And (MODALITY1 = "CG" Or MODALITY1 = "CN") Then
    CGCNREPORT
ElseIf DEPT1 = "LZZ" And (MODALITY1 = "MM" Or MODALITY1 = "MR" Or MODALITY1 = "CR" Or MODALITY1 = "CT" Or MODALITY1 = "US") Then
    RISREPORT
End If
End Sub

Private Sub cmb_procedure_Click()
On Error GoTo MSG:
cmb_procedurecode.ListIndex = cmb_procedure.ListIndex
If cmb_procedurecode.Text = "F104" Or cmb_procedurecode.Text = "F105" Then
    cmb_outblock.Visible = True
    cmb_outslide.Visible = True
    txt_outspecimen.Visible = True
    txt_outspecimen.Enabled = True
    cmb_outblock.Enabled = True
    cmb_outslide.Enabled = True
    lbl_outsideblock.Visible = True
    lbl_outside.Visible = True
    lbl_outsideslide.Visible = True
    
Else
    txt_outspecimen.Enabled = False
    cmb_outblock.Enabled = False
    cmb_outslide.Enabled = False
    lbl_outsideblock.Visible = False
    lbl_outside.Visible = False
    lbl_outsideslide.Visible = False
    cmb_outblock.Visible = False
    cmb_outslide.Visible = False
    txt_outspecimen.Visible = False
End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub cmb_procedure_GotFocus()
'cmb_procedure.Width = 5000
End Sub

Private Sub cmb_procedure_LostFocus()
'cmb_procedure.Width = 3555

End Sub

Private Sub cmb_procedurecode_Click()
On Error GoTo MSG:
cmb_procedure.ListIndex = cmb_procedurecode.ListIndex
If cmb_procedurecode.Text = "F104" Or cmb_procedurecode.Text = "F105" Then
    txt_outspecimen.Enabled = True
    cmb_outblock.Enabled = True
    cmb_outslide.Enabled = True
Else
    txt_outspecimen.Enabled = False
    cmb_outblock.Enabled = False
    cmb_outslide.Enabled = False
End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub cmb_prodis_Click()
'cmb_dept.Enabled = True
'cmb_modality.Enabled = True

End Sub

Private Sub cmb_reqno_Click()
  VIEWRETRIVE

End Sub

Private Sub cmb_vpp_GotFocus()
On Error GoTo MSG:
Dim ind_cbodept As Integer
cmb_vpp.clear
cmb_vpp.AddItem "VPP"
cmb_vpp.AddItem "DER"
cmb_vpp.AddItem "BAR"
cmb_vpp.AddItem "DrOp"
'cmb_vpp.AddItem "CHIP"

If Trim(cmb_vpp) = "" Then
    For ind_cbodept = 0 To (cmb_vpp.ListCount - 1)
        If Trim(cmb_vpp.List(ind_cbodept)) = "NONE" Then
            cmb_vpp.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Private Sub CMD_ADDNEW_Click()
Data1.Recordset.AddNew

End Sub

Private Sub cmd_back_Click()
If IHFLAG = False Then
Me.Hide
btnenabl ("INI")
FLDENABL ("DIS")
RCLEAR
IHCLEAR
FACLEAR
CGCLEAR
MBCLEAR
EDCLEAR
NoMemo = False
txt_indctinvstg.Text = ""
surgicalPathology.Visible = False
immunohistochemistry.Visible = False
fnac.Visible = False
CYTOLOGY.Visible = False
INNC = False
tmflag = False
ELFLAG = False
BCFLAG = False
ECNEWFLAG = False
ECMODFLAG = False
lbl_req.Left = 4400
lbl_req.Caption = "REQUISITION"
pathology.Caption = "Requisition"

If otlnk = True Then
    Exit Sub
End If

DIS_MAIN.show
Else
If FASPFLAG = True Then
    Me.Hide
    FLDENABL ("DIS")
    btnenabl ("INI")
    FASPFLAG = False
    IHFLAG = False
Else
    Me.Hide
    SPPATHREPORT.txt_immhischereqno.Text = txt_ihreqno.Text
    FLDENABL ("DIS")
    btnenabl ("INI")
    IHFLAG = False
End If
End If
End Sub

Private Sub cmd_cancel_Click()
btnenabl ("INI")
FLDENABL ("INI")
lbl_reqno.Visible = False
lbl_pathno.Visible = False
cmb_reqno.Visible = False
txt_pathno.Visible = False
viewflag = False
NEWFLAG = False
NEWFLAG_BCD = False
VIEWFLAG_BCD = False
PRINTNEW = False
NoMemo = False
SPCLEAR
RCLEAR
IHCLEAR
FACLEAR
CGCLEAR
MBCLEAR
EDCLEAR
CYTOLOGY.Visible = False
surgicalPathology.Visible = False
immunohistochemistry.Visible = False
frm_hematopathology.Visible = False
INNC = False
tmflag = False
BCFLAG = False
ELFLAG = False
ECNEWFLAG = False
PRINTNEW = False
bceltm.Visible = False
fnac.Visible = False
ECMODFLAG = False
Microbiology.Visible = False
txt_indctinvstg.Text = ""
lbl_req.Left = 4400
lbl_req.Caption = "REQUISITION"
pathology.Caption = "Requisition"
TEXT1.Visible = False
Label1.Visible = False
'' SANDHYA PROTOCOL PATIENT
Label12.Enabled = False
Label12.Visible = False
CMBIRB.Enabled = False
CMBIRB.Visible = False
'' END OF PROTOCOL PATIENT
End Sub

Private Sub cmd_edindication_Click()
cardiologyind.show
cardiologyind.chk_Surgery.Value = 1
If Trim(txt_modcode.Text) = "2D" Then
 cardiologyind.echoi.Visible = True
 cardiologyind.pet.Visible = False
 cardiologyind.sti.Visible = False
 
End If
If Trim(txt_modcode.Text) = "PF" Then
 cardiologyind.echoi.Visible = False
 cardiologyind.pet.Visible = True
 cardiologyind.sti.Visible = False
End If
If Trim(txt_modcode.Text) = "ST" Then
 cardiologyind.echoi.Visible = False
 cardiologyind.pet.Visible = False
 cardiologyind.sti.Visible = True
End If
If ECMODFLAG = True Then
    ECVIEW
End If
End Sub

Private Sub cmd_modify_Click()
On Error GoTo MSG:
Dim str1 As String
Dim str2 As String
Dim STR3 As String
MODIFYFLAG = True
'NM_FLG = True
'If txt_modcode.Text = "BM" Then
'    btnenabl ("NEW")
'    SPENABLE ("MOD")
'End If
If txt_modcode.Text = "SP" Then
'        MsgBox "There are no fields for Modification"
        btnenabl ("NEW")
        SPENABLE ("MOD")
        '******
                      ' SITES FILLING IN THE COMBOBOX
                  If SELMA301FLAG = False Then
                      Set SELMA301_RS = SELMA301_QR.OpenResultset(1, 3)
                      SELMA301FLAG = True
                  Else
                      Set SELMA301_RS = SELMA301_QR.OpenResultset(1, 3)
                      SELMA301_RS.MoreResults
                      SELMA301_RS.Requery
                  End If
                  
                  If SELMA301_RS.RowCount > 0 Then
                      If cmb_sites.ListCount = 0 And Len(Trim(cmb_sites.Text)) = 0 Then
                        cmb_sites.clear
                        While Not SELMA301_RS.EOF
                        If Trim(SELMA301_RS("sitedesc")) <> "" Then
                            cmb_sites.AddItem Trim(SELMA301_RS("SITEDESC")) & "$" & Trim(SELMA301_RS("SITECD"))
                        End If
                        SELMA301_RS.MoveNext
                        Wend
                      Else
                        str = cmb_sites.Text
                        cmb_sites.clear
                        While Not SELMA301_RS.EOF
                        If Trim(SELMA301_RS("sitedesc")) <> "" Then
                            cmb_sites.AddItem Trim(SELMA301_RS("SITEDESC")) & "$" & Trim(SELMA301_RS("SITECD"))
                        End If
                        SELMA301_RS.MoveNext
                        Wend
                        cmb_sites.Text = str
                      End If
                  End If
                  ' SITES FILLING IN THE COMBOBOX
                  ' LATERALITY FILLING IN THE COMBOBOX
                  If SELMA10122FLAG = False Then
                      Set SELMA10122_RS = SELMA10122_QR.OpenResultset(1, 3)
                      SELMA10122FLAG = True
                  Else
                  
                      SELMA10122_RS.MoreResults
                      SELMA10122_RS.Requery
                  End If
                  If SELMA10122_RS.RowCount > 0 Then
                      If cmb_laterality1.ListCount = 0 And Len(Trim(cmb_laterality1.Text)) = 0 Then
                        cmb_laterality1.clear
                        While Not SELMA10122_RS.EOF
                          cmb_laterality1.AddItem Trim(SELMA10122_RS("ADJ_PRVDIA"))
                          SELMA10122_RS.MoveNext
                        Wend
                       Else
                        str = ""
                        str = cmb_laterality1.Text
                        While Not SELMA10122_RS.EOF
                          cmb_laterality1.AddItem Trim(SELMA10122_RS("ADJ_PRVDIA"))
                          SELMA10122_RS.MoveNext
                        Wend
                        cmb_laterality1.Text = str
                        str = ""
                       End If
                  End If
                  If SELMA10122_RS.RowCount > 0 Then
                      If cmb_laterality.ListCount = 0 And Len(Trim(cmb_laterality.Text)) = 0 Then
                      cmb_laterality.clear
                      While Not SELMA10122_RS.EOF
                      
                      cmb_laterality.AddItem Trim(SELMA10122_RS("ADJ_PRVDIA"))
                      SELMA10122_RS.MoveNext
                      Wend
                      Else
                      str = ""
                      str = cmb_laterality.Text
                      While Not SELMA10122_RS.EOF
                        cmb_laterality.AddItem Trim(SELMA10122_RS("ADJ_PRVDIA"))
                        SELMA10122_RS.MoveNext
                      Wend
                      cmb_laterality.Text = str
                      End If
                  End If
                  
                  
                  
                  
                  ' LATERALITY FILLING IN THE COMBOBOX
                If cmb_outblock.ListCount = 0 And Len(Trim(cmb_outblock.Text)) = 0 Then
                    cmb_outblock.clear
                    For I = 0 To 5
                        cmb_outblock.AddItem I
                        
                    Next I
                Else
                    For I = 0 To 5
                        cmb_outblock.AddItem I
                        
                    Next I
                End If
                If cmb_outslide.ListCount = 0 And Len(Trim(cmb_outslide.Text)) = 0 Then
                    cmb_outslide.clear
                    For I = 0 To 5
                     
                     cmb_outslide.AddItem I
                    Next I
                Else
                    For I = 0 To 5
                     
                     cmb_outslide.AddItem I
                    Next I
                End If
                txt_spreqno.Enabled = False
                txt_sppathno.Enabled = False
                ' FILLING THE PREVIOUSLY INVESTIGATED REQ NO.
                 
'                 If SELRREQUISITNFLAG = False Then
'                      SELRREQUISITN_QR(0) = Msk_caseno.FormattedText
'                      Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
'                      SELRREQUISITNFLAG = True
'                 Else
'
'
'                      SELRREQUISITN_QR(0) = Msk_caseno.FormattedText
'                      SELRREQUISITN_RS.MoreResults
'                      SELRREQUISITN_RS.Requery
'
'                  End If
'                  If SELRREQUISITN_RS.RowCount > 0 Then
'                      cmb_preinvest.clear
'                      While Not SELRREQUISITN_RS.EOF
'                          cmb_preinvest.AddItem SELRREQUISITN_RS("REQNNO")
'                          SELRREQUISITN_RS.MoveNext
'                      Wend
'                      cmb_preinvest.ListIndex = 0
'                  End If
''commented by mono
'                If seLdreqrsflag = False Then
'                    SELDREQQR(0) = Msk_caseno.FormattedText
'                    Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
'                    seLdreqrsflag = True
'                Else
'                    SELDREQQR(0) = Msk_caseno.FormattedText
'                    SELDREQRS.MoreResults
'                    SELDREQRS.Requery
'                End If
'                If SELDREQRS.RowCount > 0 Then
'                    cmb_preinvest.clear
'                    While Not SELDREQRS.EOF
'                        cmb_preinvest.AddItem SELDREQRS("PATHNO") & "$" & SELDREQRS("REQNO")
'                        SELDREQRS.MoveNext
'                    Wend
'                    cmb_preinvest.ListIndex = 0
'                End If
                ' FILLING THE PREVIOUSLY INVESTIGATED REQ NO.
                'FILLING THE OPERATIVE PROCEDURES
                If SELMA101FLAG1 = False Then
                  Set SELMA101RS = selMA101QR.OpenResultset(1, 3)
                  SELMA101FLAG1 = True
                Else
                  SELMA101RS.MoreResults
                  SELMA101RS.Requery
                End If
'                If SELMA101RS.RowCount > 0 Then
'                  txt_opprocedure.clear
'                  While Not SELMA101RS.EOF
'                      txt_opprocedure.AddItem Trim(SELMA101RS("SER_DET"))
'                      SELMA101RS.MoveNext
'                  Wend
'                End If
                'FILLING THE OPERATIVE PROCEDURES
                str = cmb_spfnac.Text
                cmb_spfnac.clear
                cmb_spfnac.AddItem "YES"
                cmb_spfnac.AddItem "NO"
                cmb_spfnac.Text = str
       
        '******
 End If
 If txt_modcode.Text = "IH" Then
    btnenabl ("NEW")
    IHENABL ("MOD")
    cmb_refdoc.Enabled = True
 End If
 If txt_modcode.Text = "BP" Then
    btnenabl ("NEW")
    IHENABL ("MOD")
    cmb_refdoc.Enabled = True
 End If
 If txt_modcode.Text = "HT" Then
    btnenabl ("NEW")
    IHENABL ("MOD")
    cmb_refdoc.Enabled = True
 End If
 If txt_modcode.Text = "HL" Then
    btnenabl ("NEW")
    IHENABL ("MOD")
    cmb_refdoc.Enabled = True
 End If
 If txt_modcode.Text = "FA" Then
    btnenabl ("NEW")
    FAENABL ("MOD")
    cmb_refdoc.Enabled = True
    ' SITES FILLING IN THE COMBOBOX
    If SELMA301FLAG = False Then
        Set SELMA301_RS = SELMA301_QR.OpenResultset(1, 3)
        SELMA301FLAG = True
    Else
        Set SELMA301_RS = SELMA301_QR.OpenResultset(1, 3)
        SELMA301_RS.MoreResults
        SELMA301_RS.Requery
    End If
    
    If SELMA301_RS.RowCount > 0 Then
        If cmb_fnsite.ListCount = 0 And Len(Trim(cmb_fnsite.Text)) = 0 Then
                        cmb_fnsite.clear
                        While Not SELMA301_RS.EOF
                        If Trim(SELMA301_RS("sitedesc")) <> "" Then
                            cmb_fnsite.AddItem Trim(SELMA301_RS("SITEDESC")) & "$" & Trim(SELMA301_RS("SITECD"))
                        End If
                        SELMA301_RS.MoveNext
                        Wend
                      Else
                        str = cmb_fnsite.Text
                        cmb_fnsite.clear
                        While Not SELMA301_RS.EOF
                        If Trim(SELMA301_RS("sitedesc")) <> "" Then
                            cmb_fnsite.AddItem Trim(SELMA301_RS("SITEDESC")) & "$" & Trim(SELMA301_RS("SITECD"))
                        End If
                        SELMA301_RS.MoveNext
                        Wend
                        cmb_fnsite.Text = str
                      End If
    End If
    ' SITES FILLING IN THE COMBOBOX
    ' LATERALITY FILLING IN THE COMBOBOX
    If SELMA10122FLAG = False Then
        Set SELMA10122_RS = SELMA10122_QR.OpenResultset(1, 3)
        SELMA10122FLAG = True
    Else
    
        SELMA10122_RS.MoreResults
        SELMA10122_RS.Requery
    End If
    If SELMA10122_RS.RowCount > 0 Then
           If cmb_fnlaterality.ListCount = 0 And Len(Trim(cmb_fnlaterality.Text)) = 0 Then
                cmb_fnlaterality.clear
                While Not SELMA10122_RS.EOF
                  cmb_fnlaterality.AddItem Trim(SELMA10122_RS("ADJ_PRVDIA"))
                  SELMA10122_RS.MoveNext
                Wend
            Else
                str = ""
                str = cmb_fnlaterality.Text
                While Not SELMA10122_RS.EOF
                  cmb_fnlaterality.AddItem Trim(SELMA10122_RS("ADJ_PRVDIA"))
                  SELMA10122_RS.MoveNext
                Wend
                cmb_fnlaterality.Text = str
                str = ""
            End If
    End If
    If SELMA10122_RS.RowCount > 0 Then
        If cmb_fnlaterality1.ListCount = 0 And Len(Trim(cmb_fnlaterality1.Text)) = 0 Then
                cmb_fnlaterality1.clear
                SELMA10122_RS.MoveFirst
                While Not SELMA10122_RS.EOF
                
                cmb_fnlaterality1.AddItem Trim(SELMA10122_RS("ADJ_PRVDIA"))
                SELMA10122_RS.MoveNext
                Wend
        Else
                str = ""
                str = cmb_fnlaterality1.Text
                SELMA10122_RS.MoveFirst
                While Not SELMA10122_RS.EOF
                  cmb_fnlaterality1.AddItem Trim(SELMA10122_RS("ADJ_PRVDIA"))
                  SELMA10122_RS.MoveNext
                Wend
                cmb_fnlaterality1.Text = str
        End If
    End If
    ' LATERALITY FILLING IN THE COMBOBOX
    ' FILLING THE PREVIOUSLY INVESTIGATED REQ NO.
                If seLdreqrsflag = False Then
                    SELDREQQR(0) = Msk_caseno.FormattedText
                    Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                    seLdreqrsflag = True
                Else
                    SELDREQQR(0) = Msk_caseno.FormattedText
                    SELDREQRS.MoreResults
                    SELDREQRS.Requery
                End If
                If SELDREQRS.RowCount > 0 Then
                    cmb_preinvest.clear
                    While Not SELDREQRS.EOF
                        cmb_fnpreinvest.AddItem SELDREQRS("PATHNO") & "$" & SELDREQRS("REQNO")
                        SELDREQRS.MoveNext
                    Wend
                    cmb_fnpreinvest.ListIndex = 0
                End If

  ' FILLING THE PREVIOUSLY INVESTIGATED REQ NO.
'  CMB_CELLBLOCK.AddItem "YES"
'  CMB_CELLBLOCK.AddItem "NO"
  'FILLING THE OUTSIDESLIDES
  str1 = TXT_FNSLIDEDRY.Text
  str2 = TXT_FNOSLIDEWET.Text
  STR3 = cmb_outsideslidesfn.Text
  TXT_FNSLIDEDRY.clear
  TXT_FNOSLIDEWET.clear
  cmb_outsideslidesfn.clear
  For I = 0 To 20
    TXT_FNSLIDEDRY.AddItem I
    TXT_FNOSLIDEWET.AddItem I
    cmb_outsideslidesfn.AddItem I
  Next I
  TXT_FNSLIDEDRY.Text = str1
  TXT_FNOSLIDEWET.Text = str2
  cmb_outsideslidesfn.Text = STR3
  'FNAC PERFORMED BY
  If SELDENTERFLAG = False Then
        Set SELDENTERRS = SELDENTERQR.OpenResultset(1, 3)
        SELDENTERFLAG = True
    Else
        SELDENTERRS.MoreResults
        SELDENTERRS.Requery
    End If
    If SELDENTERRS.RowCount > 0 Then
        str1 = ""
        str1 = cmb_fnac.Text
        cmb_fnac.clear
        While Not SELDENTERRS.EOF
            cmb_fnac.AddItem Trim(SELDENTERRS("CCNAME"))
            SELDENTERRS.MoveNext
        Wend
        cmb_fnac.Text = str1
    End If
  
  
End If
'**** CG OR CN
If Trim(txt_modcode.Text) = "CN" Or Trim(txt_modcode.Text) = "CG" Or Trim(txt_modcode.Text) = "PG" Or Trim(txt_modcode.Text) = "PN" Then
    btnenabl ("NEW")
    CGENABL ("MOD")
    cmb_refdoc.Enabled = True
    ' SITES FILLING IN THE COMBOBOX
    If txt_modcode.Text = "CN" Then
    If SELMA301FLAG = False Then
        Set SELMA301_RS = SELMA301_QR.OpenResultset(1, 3)
        SELMA301FLAG = True
    Else
        Set SELMA301_RS = SELMA301_QR.OpenResultset(1, 3)
        SELMA301_RS.MoreResults
        SELMA301_RS.Requery
    End If
    
    If SELMA301_RS.RowCount > 0 Then
    
                If cmb_cnsite.ListCount = 0 And Len(Trim(cmb_cnsite.Text)) = 0 Then
                        cmb_fnsite.clear
                        While Not SELMA301_RS.EOF
                        If Trim(SELMA301_RS("sitedesc")) <> "" Then
                            cmb_nsite.AddItem Trim(SELMA301_RS("SITEDESC")) & "$" & Trim(SELMA301_RS("SITECD"))
                        End If
                        SELMA301_RS.MoveNext
                        Wend
                      Else
                        str = cmb_cnsite.Text
                        cmb_cnsite.clear
                        While Not SELMA301_RS.EOF
                        If Trim(SELMA301_RS("sitedesc")) <> "" Then
                            cmb_cnsite.AddItem Trim(SELMA301_RS("SITEDESC")) & "$" & Trim(SELMA301_RS("SITECD"))
                        End If
                        SELMA301_RS.MoveNext
                        Wend
                        cmb_cnsite.Text = str
                      End If
    End If
    End If
    If txt_modcode.Text = "CG" Then
        str1 = cmb_cnsite.Text
        cmb_cnsite.clear
        cmb_cnsite.AddItem "VULVA"
        cmb_cnsite.AddItem "VAGINA"
        cmb_cnsite.AddItem "CERVIX"
        cmb_cnsite.AddItem "ENDOMETRIUM"
        cmb_cnsite.AddItem "PERITIONEAL FLUID"
        cmb_cnsite.AddItem "INGUINAL NODE"
        cmb_cnsite.AddItem "OTHERS"
        cmb_cnsite.Text = str1
    End If
    ' SITES FILLING IN THE COMBOBOX
    ' LATERALITY FILLING IN THE COMBOBOX
    If SELMA10122FLAG = False Then
        Set SELMA10122_RS = SELMA10122_QR.OpenResultset(1, 3)
        SELMA10122FLAG = True
    Else
    
        SELMA10122_RS.MoreResults
        SELMA10122_RS.Requery
    End If
    If SELMA10122_RS.RowCount > 0 Then
           
                str = ""
                str = cmb_claterality.Text
                cmb_claterality.clear
                While Not SELMA10122_RS.EOF
                  cmb_claterality.AddItem Trim(SELMA10122_RS("ADJ_PRVDIA"))
                  SELMA10122_RS.MoveNext
                Wend
                cmb_claterality.Text = str
                str = ""
           
    End If
    ' LATERALITY FILLING IN THE COMBOBOX
    ' FILLING THE PREVIOUSLY INVESTIGATED REQ NO.
                If seLdreqrsflag = False Then
                    SELDREQQR(0) = Msk_caseno.FormattedText
                    Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                    seLdreqrsflag = True
                Else
                    SELDREQQR(0) = Msk_caseno.FormattedText
                    SELDREQRS.MoreResults
                    SELDREQRS.Requery
                End If
                If SELDREQRS.RowCount > 0 Then
                    'cmb_cpreinv.clear
                    While Not SELDREQRS.EOF
                        'cmb_cpreinv.AddItem SELDREQRS("PATHNO") & "$" & SELDREQRS("REQNO")
                        'SELDREQRS.MoveNext
                        'REQNNO,REQNDT,REQNSTAT,CASENO,PREPORTDT,PATHNO
                    Wend
                    cmb_cpreinv.ListIndex = 0
                End If

  ' FILLING THE PREVIOUSLY INVESTIGATED REQ NO.
        str1 = cmb_cnoslides.Text
        str2 = cmb_cyoutspecimen.Text
        STR3 = cmb_coutslides.Text
        cmb_cnoslides.clear
        'cmb_cyoutspecimen.clear
        cmb_coutslides.clear
                 For I = 0 To 20
                    cmb_cnoslides.AddItem I
                    'cmb_cyoutspecimen.AddItem i
                    cmb_coutslides.AddItem I
                Next I
        cmb_cnoslides.Text = str1
        'cmb_cyoutspecimen.Text = STR2
        cmb_coutslides.Text = STR3
    
    
End If
'**** CG OR CN
'**** MICROBIOLOGY
If Trim(txt_modcode.Text) = "BA" Or Trim(txt_modcode.Text) = "MB" Or Trim(txt_modcode.Text) = "SE" Or Trim(txt_modcode.Text) = "MC" Or Trim(txt_modcode.Text) = "CM" Or Trim(txt_modcode.Text) = "MO" Or Trim(txt_modcode.Text) = "XX" Then
    MBENABL ("MOD")
    btnenabl ("NEW")
    str = ""
    str = cmb_mboutslide
    cmb_mboutslide.clear
    For I = 0 To 10
        cmb_mboutslide.AddItem I
    Next I
    cmb_mboutslide.Text = str
    str = ""
    txt_indctinvstg.Enabled = True
    cmb_refdoc.Enabled = True
End If
    
    
'**** MICROBIOLOGY
'**** ENDOSCOPY
If Trim(txt_modcode.Text) = "ED" Then
    Dim a As Integer
    Dim a1 As String
    Dim b As Integer
    Dim b1 As String
    Dim c As Integer
    Dim C1 As String
    Dim D As Integer
    Dim D2 As String
    Dim e As Integer
    Dim e1 As String
    EDENABL ("MOD")
    btnenabl ("NEW")
            a = InStr(list_edts.Text, "Biopsy")
            If a > 0 Then
                a1 = Mid(list_edts, a, 6)
            End If
            b = InStr(list_edts.Text, "Brushing")
            If b > 0 Then
                b1 = Mid(list_edts, b, 8)
            End If
            c = InStr(list_edts.Text, "FNA")
            If c > 0 Then
                C1 = Mid(list_edts, c, 3)
            End If
            D = InStr(list_edts.Text, "Culture")
            If D > 0 Then
                D2 = Mid(list_edts, D, 7)
            End If
            
            e = InStr(list_edts.Text, "Aspiration")
            If e > 0 Then
                e1 = Mid(list_edts, e, 10)
            End If
            list_edts.clear
            list_edts.AddItem "Biopsy"
            list_edts.AddItem "Brushing"
            list_edts.AddItem "FNA"
            list_edts.AddItem "Culture"
            list_edts.AddItem "Aspiration"
            
            If Len(a1) > 0 Then
                list_edts.Selected(0) = True
            End If
            If Len(b1) > 0 Then
                list_edts.Selected(1) = True
            End If
            If Len(C1) > 0 Then
                list_edts.Selected(2) = True
            End If
            If Len(D1) > 0 Then
                list_edts.Selected(3) = True
            End If
            If Len(e1) > 0 Then
                list_edts.Selected(4) = True
            End If
            str1 = cmb_edprocedureheading.Text
            cmb_edprocedureheading.clear
            cmb_edprocedureheading.AddItem "UPPER G.I. ENDOSCOPY"
            cmb_edprocedureheading.AddItem "SIGMOIDOSCOPY"
            cmb_edprocedureheading.AddItem "COLONOSCOPY"
            cmb_edprocedureheading.AddItem "ERCP"
            cmb_edprocedureheading.AddItem "ENDOSCOPIC ULTRASONOGRAPHY"
            cmb_edprocedureheading.Text = str1
            str1 = ""

End If
'**** ENDOSCOPY
'**** CARDIOLOGY
If Trim(txt_modcode.Text) = "2D" Or Trim(txt_modcode.Text) = "EK" Or Trim(txt_modcode.Text) = "PF" Or Trim(txt_modcode.Text) = "ST" Then
        btnenabl ("NEW")
            
            cmb_hpprocedure.Enabled = True
            cmb_hpprocedurecode.Enabled = True
            If Trim(txt_modcode.Text) <> "EK" Then
            cmd_edindication.Visible = True
            cmd_edindication.Enabled = True
            ECMODFLAG = True
            End If
            
            If SELRINVESTFLAG = False Then
        SELRINVEST_QR(0) = Trim(txt_modcode.Text)
        Set SELRINVEST_RS = SELRINVEST_QR.OpenResultset(1, 3)
        SELRINVESTFLAG = True
    Else
        SELRINVEST_QR(0) = Trim(txt_modcode.Text)
        SELRINVEST_RS.MoreResults
        SELRINVEST_RS.Requery
    End If
    
    If SELRINVEST_RS.RowCount > 0 Then
        str1 = cmb_hpprocedure.Text
        str2 = cmb_hpprocedurecode.Text
        cmb_hpprocedure.clear
        cmb_hpprocedurecode.clear
        While Not SELRINVEST_RS.EOF
            cmb_hpprocedurecode.AddItem Trim(SELRINVEST_RS("ser_code"))
            'cbo_proccode.Text = SELRINVEST_RS("ser_code")
            If SELMA101FLAG = False Then
                SELMA101_QR(0) = Trim(SELRINVEST_RS("SER_CODE"))
                Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                SELMA101FLAG = True
             Else
                SELMA101_QR(0) = SELRINVEST_RS("SER_CODE")
                SELMA101_RS.MoreResults
                SELMA101_RS.Requery
            End If
            If SELMA101_RS.RowCount > 0 Then
                While Not SELMA101_RS.EOF
                cmb_hpprocedure.AddItem SELMA101_RS("SER_DET")
                'CMB_CGPROCEDURE.Text = SELMA101_RS("SER_DET")
                SELMA101_RS.MoveNext
                Wend
            End If
            SELRINVEST_RS.MoveNext
        Wend
        
          cmb_hpprocedure.Text = str1
          cmb_hpprocedurecode.Text = str2
          str1 = ""
          str2 = ""
    End If
        

End If

'**** CARDIOLOGY
 Exit Sub
MSG:
 MsgBox Err.Description
End Sub

Private Sub cmd_new_Click()

'On Error GoTo MSG:

Dim I As Integer
NoMemo = False
'V_PROCE = False
btnenabl ("NEW")
FLDENABL ("NEW")
SPCLEAR
RCLEAR
IHCLEAR
'********** Added by Shraddha on 26/03/09 as told by Sandhya m'am **********
CMB_TEXT.clear
'***************************************************************************
txt_clinhist.Text = "-"
txt_indctinvstg.Text = "-"

cmb_prodis.Text = "-"
If DATETIMEFLAG = False Then
   DATETIME
   DATETIMEFLAG = True
Else
    DATETIME
End If
txt_date.Text = DATE11
txt_time.Text = time1

lbl_reqno.Visible = False
lbl_pathno.Visible = False
cmb_reqno.Visible = False
txt_pathno.Visible = False
TEXT1.Visible = False
Label1.Visible = False
'msk_caseno.SetFocus
NEWFLAG = True
NEWFLAG_BCD = True
PRINTNEW = True
viewflag = False
meterial = ""
'' SANDHYA PROTOCOL
CMBIRB.Text = " "
CMBIRB.Visible = False
MIRBNO = ""
'' END PROTO

ECNEWFLAG = False
ECMODFLAG = False
If opt_patient.Value = True Then
    lbl_caseno.Caption = "Case No."
    Msk_caseno.Mask = "??>/#####"
    LBL_DEPENDENT.Visible = False
    cmb_dependents.Visible = False
     Msk_caseno.Enabled = True
   ' msk_caseno.SetFocus
End If
If opt_staff.Value = True Then
    lbl_caseno.Caption = "  CC No."
    Msk_caseno.Mask = "########"
    LBL_DEPENDENT.Visible = True
    cmb_dependents.Visible = True
    Msk_caseno.Enabled = True
   ' msk_caseno.SetFocus
End If
'opt_patient.Value = True
'999
'Data1.Refresh
'I = Data1.Recordset.RecordCount
'If I <> 0 Then
 '   Data1.Recordset.MoveLast
  '  Data1.Recordset.MoveFirst 999
  

'Data1.Recordset.MoveFirst

'I = Data1.Recordset.RecordCount
'For K = 0 To I - 1

'With Data1.Recordset
'    .delete
'    .MoveNext
'    If .EOF Then Exit Sub
'End With
'Next
'End If
'999
opt_patient.Value = True
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub cmd_print_Click()
Dim I As Integer
Dim str1 As String
Dim prepath As String
Dim k1 As Integer
Dim dept As String, MODALITY As String
'On Error GoTo MSG:
'  PATCHING FOR NEW MODE TO CHECK WEATHER PAID STATUS ARE MEMO NEED TO BE PREPARED
   'If PRINTNEW = True Then
If lbl_icureq.Visible = True Then ' CHECKING ICU not for other requisitn


           If SELRISSTAT_FLAG = False Then
               SELRISSTAT_QR(0) = req_chglen(txt_bcreqno.Text)
                 'SELRISSTAT_QR(0) = txt_bcreqno.Text
                 Set SELRISSTAT_RS = SELRISSTAT_QR.OpenResultset(1, 3)
                 SELRISSTAT_FLAG = True
            Else
                 SELRISSTAT_QR(0) = req_chglen(txt_bcreqno.Text)
                 'SELRISSTAT_QR(0) = txt_bcreqno.Text
                 SELRISSTAT_RS.MoreResults
                 SELRISSTAT_RS.Requery
            End If
   
           If Not SELRISSTAT_RS.EOF Then
              If (SELRISSTAT_RS("REQNSTAT") = "P") Or (SELRISSTAT_RS("PAT_CATG") = "NC") Or (SELRISSTAT_RS("PAT_CATG") = "SF") Then
              
              Else
                
                If NEWFLAG = True Then
                    If IsNull(SELRISSTAT_RS("memono")) And ((Trim(txt_catgory.Text) <> "NC") Or (opt_staff.Value = False) Or (SELRISSTAT_RS("PAT_CATG") = "SF")) Then
                        'MsgBox " Please Prepare Memo !!!"
                        PRINTNEW = False
                        Exit Sub
                    End If
                End If
                dept = Right(Trim(cmb_dept.Text), 3)
                MODALITY = Right(Trim(cmb_modality.Text), 2)
                
                If IsNull(SELRISSTAT_RS("memono")) And ((Trim(dept) = "FEC") Or (Trim(dept) = "CBZ") Or (Trim(dept) = "FZZ" And MODALITY <> "SP")) Then
                    MsgBox " Please Prepare Memo !!!"
                    PRINTNEW = False
                    Exit Sub
                End If
                
              End If
              
            End If
End If  ' CHECKING ICU not for other requisitn over
    
    'End If
 '  PATCHING FOR NEW MODE TO CHECK WEATHER PAID STATUS ARE MEMO NEED TO BE PREPARED

 
'' COMMON PRINTING
'Dim DeviceName As String
'Dim CURPRT As String
'
'CURPRT = UCase(Printer.DeviceName)
''DeviceName = "Epson LQ-1050+"
'DeviceName = "COMMON"
'
'If DeviceName <> CURPRT Then
''If PRT_FLAG_STS = False Then
'   'XA = List1.List(List1.ListIndex)
'
'
'    If CSETPRINTER.SetPrinterAsDefault(DeviceName) Then
'        'sMsg = DeviceName & " has successfully been set as the default printer."
'     Else
'        MsgBox " " & DeviceName & " has failed to be set as the default printer."
'        Exit Sub
'    End If
'''CommonDialog1.flags = &H40
'
''CommonDialog1.Action = 5
''End If
'End If
'
'CURPRT = UCase(Printer.DeviceName)
'
'If DeviceName <> CURPRT Then
'MsgBox "PLEASE CHANGE DEFULAT PRINTER TO " & DeviceName & " ", vbInformation
''CommonDialog1.flags = &H40
''CommonDialog1.Action = 5
'Exit Sub
'Else
''PRT_FLAG_STS = True
'End If
''COMMON PRINTING
    For I = 0 To 25
    CrystalReport1.Formulas(I) = ""
    Next I

        If seLdreqrsflag = False Then
            SELDREQQR(0) = Trim(Msk_caseno.FormattedText)
            Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
            seLdreqrsflag = True
        Else
            SELDREQQR(0) = Trim(Msk_caseno.FormattedText)
            SELDREQRS.MoreResults
            SELDREQRS.Requery
        End If
        prepath = ""
        If SELDREQRS.RowCount > 0 And Len(SELDREQRS("pathno")) = 9 Then
            prepath = ""
            SELDREQRS.MoveNext
            While Not SELDREQRS.EOF
                prepath = prepath & SELDREQRS("PATHNO") & " "
                SELDREQRS.MoveNext
            Wend
            prepath = Trim(prepath)
            prepath = Left(prepath, 100)
        End If
        I = Len(Trim(prepath))
        If I = 0 Then
            prepath = ""
        Else
            prepath = prepath
            'prepath = Mid(prepath, 1, i - 1)
        End If
        str = ipadd & " " & Format(DATE11, "dd/mm/yyyy") & " " & time1
       'MsgBox STR
If opt_patient.Value = True Then
        If SELDTMNO_FLAG = False Then
            SELDTMNO_QR(0) = Trim(Msk_caseno.FormattedText)
            Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
            SELDTMNO_FLAG = True
        Else
           SELDTMNO_QR(0) = Trim(Msk_caseno.FormattedText)
           SELDTMNO_RS.MoreResults
           SELDTMNO_RS.Requery
        End If
        If viewflag = True Then
            If SELRISSTAT_FLAG = False Then
                 SELRISSTAT_QR(0) = req_chglen(cmb_reqno.Text)
                 Set SELRISSTAT_RS = SELRISSTAT_QR.OpenResultset(1, 3)
                 SELRISSTAT_FLAG = True
            Else
                 SELRISSTAT_QR(0) = req_chglen(cmb_reqno.Text)
                 SELRISSTAT_RS.MoreResults
                 SELRISSTAT_RS.Requery
            End If
            If Not SELRISSTAT_RS.EOF Then
               If Trim(txt_catgory.Text) = "NC" Then
               
               Else
'                If IsNull(SELRISSTAT_RS("memono")) Then
'                    MsgBox " Please Prepare Memo !!!"
'                    Exit Sub
'                End If
               End If
            End If
            CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
               '***SCREENING SERVICES
            
            If txt_modcode = "SS" Then
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPSS.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & cmb_reqno.Text & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                
                
'                If Trim(prepath) = "" Then
'                    CrystalReport1.Formulas(5) = "PATH = '" & "" & " '"
'                Else
'                    CrystalReport1.Formulas(5) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
'                End If
                If txt_modcode.Text = "SS" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR SCREENING SERVICES" & " '"
                End If
                If Not SELDTMNO_RS.EOF Then
                    If IsNull(SELDTMNO_RS("DTMNO")) Then
                        CrystalReport1.Formulas(7) = "DTMNO = '" & " " & " '"
                    Else
                        CrystalReport1.Formulas(7) = "DTMNO = '" & Trim(SELDTMNO_RS("DTMNO")) & " '"
                    End If
                End If
                CrystalReport1.Formulas(8) = "ENTERBY='" & Trim(enter_by1) & "'"
                'CRYSTALREPORT1.Connect = "DSN=DISDSN;UID=" & LOGIN & ",PWD=" & LOGIN & ";DATABASE=EC_TMH"
                'CRYSTALREPORT1.Connect = "DSN=DISDSN;DATABASE=EC_TMH;UID=" & LOGIN & ",PWD=" & LOGIN & ""
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            End If
            If txt_modcode = "SP" Then
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREP.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & cmb_reqno.Text & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_sex & "/" & txt_Age & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                I = InStr(cmb_sites, "$")
                If I = 0 Then
                    CrystalReport1.Formulas(6) = "SITE = '" & cmb_sites & "'"
                Else
                    CrystalReport1.Formulas(6) = "SITE = '" & Mid(cmb_sites, 1, I - 1) & "'"
                End If
                CrystalReport1.Formulas(7) = "LATER = '" & cmb_laterality1 & "'"
                CrystalReport1.Formulas(8) = "OPPROC = '" & TXT_OPPROCEDURE & "'"
                CrystalReport1.Formulas(9) = "METAS = '" & txt_metastasis & "'"
                CrystalReport1.Formulas(10) = "LATER1 = '" & cmb_laterality & "'"
                CrystalReport1.Formulas(11) = "OUTSP = '" & txt_outspecimen & "'"
                CrystalReport1.Formulas(12) = "OUTBLK = '" & cmb_outblock & "'"
                CrystalReport1.Formulas(13) = "OUTSLID = '" & cmb_outslide & "'"
                CrystalReport1.Formulas(14) = "USRID = '" & str & "'"
                CrystalReport1.Formulas(15) = "PROV = '" & txt_indctinvstg & "'"
                If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(16) = "PATH = '" & "" & " '"
                Else
                    CrystalReport1.Formulas(16) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
                End If
                CrystalReport1.Formulas(17) = "ENTERBY='" & Trim(enter_by1) & "'"
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            End If
            If txt_modcode = "IH" Then
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPIH.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & cmb_reqno.Text & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In Patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_sex & "/" & txt_Age & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                CrystalReport1.Formulas(6) = "OUTBLK = '" & cmb_Ihsoutsideblock & "'"
                CrystalReport1.Formulas(7) = "OUTSLID = '" & cmb_Ihsoutslide & "'"
                 If seldihreportflag = False Then
                    seldihreport_qr(0) = req_chglen(cmb_reqno)
                    Set seldihreport_rs = seldihreport_qr.OpenResultset(1, 3)
                    seldihreportflag = True
                 Else
                    seldihreport_qr(0) = req_chglen(cmb_reqno.Text)
                    seldihreport_rs.MoreResults
                    seldihreport_rs.Requery
                 End If
                 If seldihreport_rs.RowCount > 0 Then
                 str1 = ""
                    While Not seldihreport_rs.EOF
                        str1 = str1 & Trim(seldihreport_rs("testorder")) & ", "
                        seldihreport_rs.MoveNext
                    Wend
                 End If
                I = Len(Trim(str1))
                If I = 0 Then
                    CrystalReport1.Formulas(8) = "TEST = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(8) = "TEST = '" & Left(Trim(str1), I - 1) & "'"
                End If
                CrystalReport1.Formulas(9) = "USRID = '" & str & "'"
                CrystalReport1.Formulas(10) = "PROV = '" & txt_indctinvstg & "'"
                If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(11) = "PATH = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(11) = "PATH = '" & "Previous Path No." & Left(Trim(prepath), 200) & "'"
                End If
                'CRYSTALREPORT1.Formulas(12) = "ENTERBY='" & Trim(Enter_by1) & "'"
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            End If
            If txt_modcode.Text = "FA" Then
                'CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPFA.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & cmb_reqno.Text & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_sex & "/" & txt_Age & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                I = InStr(cmb_sites, "$")
                If I = 0 Then
                    CrystalReport1.Formulas(6) = "SITE = '" & cmb_fnsite & "'"
                Else
                    CrystalReport1.Formulas(6) = "SITE = '" & Mid(cmb_fnsite, 1, I - 1) & "'"
                End If
                '******************** Shraddha Pandit 21/05/2009 *************************
                CrystalReport1.Formulas(20) = "NAM = '" & txt_name.Text & "'"
                CrystalReport1.Formulas(21) = "proc_name = '" & CMB_FNPROCEDURE.Text & "'"
                '*************************************************************************
                CrystalReport1.Formulas(7) = "LATER = '" & cmb_fnlaterality & "'"
                CrystalReport1.Formulas(8) = "METAS = '" & txt_fnmetastasis & "'"
                CrystalReport1.Formulas(9) = "OPPROC = '" & cmb_fnac & "'"
                CrystalReport1.Formulas(10) = "LATER1 = '" & cmb_fnlaterality1 & "'"
                CrystalReport1.Formulas(11) = "OUTSP = '" & TXT_FNSLIDEDRY & "'"
                CrystalReport1.Formulas(12) = "OUTBLK = '" & TXT_FNOSLIDEWET & "'"
                CrystalReport1.Formulas(13) = "OUTSLID = '" & cmb_outslide & "'"
                CrystalReport1.Formulas(14) = "USRID = '" & str & "'"
                CrystalReport1.Formulas(15) = "PROV = '" & txt_indctinvstg & "'"
                If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(16) = "PATH = '" & "" & " '"
                Else
                    CrystalReport1.Formulas(16) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
                End If
                CrystalReport1.Formulas(17) = "ENTERBY='" & Trim(enter_by1) & "'"
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            End If
            If txt_modcode = "CG" Or txt_modcode = "CN" Or txt_modcode = "PG" Or txt_modcode = "PN" Then
                'CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPCG.RPT"
                'CrystalReport1.Formulas(0) = "REQNO = '" & cmb_reqno.Text & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                
                '******* Shraddha Pandit 29/01/09 *******
                CrystalReport1.Formulas(20) = "Nam = '" & txt_name.Text & "'"
                CrystalReport1.Formulas(21) = "proce_name = '" & CMB_CGPROCEDURE.Text & "'"
                '****************************************
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "st = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "st = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "st = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SexAge = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_sex & "/" & txt_Age & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATH = '" & txt_pathno & "'"
                I = InStr(cmb_cnsite, "$")
                If I = 0 Then
                    CrystalReport1.Formulas(6) = "SITE = '" & cmb_cnsite & "'"
                Else
                    CrystalReport1.Formulas(6) = "SITE = '" & Mid(cmb_cnsite, 1, I - 1) & "'"
                End If
                CrystalReport1.Formulas(7) = "LATER = '" & cmb_claterality & "'"
                CrystalReport1.Formulas(8) = "OPPROC = '" & txt_trtdetail & "'"
                CrystalReport1.Formulas(9) = "METAS = '" & txt_metastasis & "'"
                CrystalReport1.Formulas(10) = "LATER1 = '" & txt_scopyfind & "'"
                CrystalReport1.Formulas(11) = "OUTSP = '" & cmb_cyoutspecimen & "'"
                CrystalReport1.Formulas(12) = "OUTBLK = '" & cmb_cnoslides & "'"
                CrystalReport1.Formulas(13) = "OUTSLID = '" & cmb_coutslides & "'"
                CrystalReport1.Formulas(14) = "USRID = '" & str & "'"
                CrystalReport1.Formulas(15) = "PROV = '" & txt_indctinvstg & "'"
                If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(16) = "PATH = '" & "" & " '"
                Else
                    CrystalReport1.Formulas(16) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
                End If
                CrystalReport1.Formulas(18) = "ENTERBY = '" & Trim(enter_by1) & "'"
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            End If
            '*** BIOCHEMISTRY
            If txt_modcode = "BC" Or txt_modcode = "EL" Or txt_modcode = "TM" Then
                'CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPBC.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & cmb_reqno.Text & "'"
                'CRYSTALREPORT1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATH = '" & txt_pathno & "'"
                
                
'                If Trim(prepath) = "" Then
'                    CrystalReport1.Formulas(5) = "PATH = '" & "" & " '"
'                Else
'                    CrystalReport1.Formulas(5) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
'                End If
                If txt_modcode.Text = "BC" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR BIOCHEMSITRY" & " '"
                End If
                If txt_modcode.Text = "EL" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR EMERGENCY LAB" & " '"
                End If
                If txt_modcode.Text = "TM" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR TUMOUR MARKER " & " '"
                End If
                CrystalReport1.Formulas(7) = "ipaddress = '" & pathwin.LocalIP & "'"
                CrystalReport1.Formulas(18) = "ENTERBY = '" & Trim(enter_by1) & "'"
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            End If
            '*** BIOCHEMSITRY
            '*** MICROBIOLOGY
            If txt_modcode = "BA" Or txt_modcode = "MB" Or txt_modcode = "SE" Or txt_modcode = "MC" Or txt_modcode = "CM" Or txt_modcode = "MO" Or txt_modcode = "XX" Then
                'CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPMB.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & cmb_reqno.Text & "'"
                'CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & Trim(txt_ward.Text) & "/" & Trim(txt_bedno.Text) & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                
                
                If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(5) = "PATH = '" & "" & " '"
                Else
                    'Next line commented and 2nd line added by Shraddha Pandit on 06/01/09
                    'CrystalReport1.Formulas(5) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
                    CrystalReport1.Formulas(5) = "PATH = '" & " Previous path No. " & Left(Replace(prepath, "'", ""), 200) & "'"
                End If
                If txt_modcode.Text = "BA" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR BACTERIOLOGY" & " '"
                End If
                If txt_modcode.Text = "MB" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR MYCOBACTERIOLOGY" & " '"
                End If
                If txt_modcode.Text = "SE" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR SEROLOGY " & " '"
                End If
                If txt_modcode.Text = "MC" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR MYCOLOGY " & " '"
                End If
                If txt_modcode.Text = "CM" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR CLINICAL MICROBIOLOGY" & " '"
                End If
                If txt_modcode.Text = "MO" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR MOLECULAR MICROBIOLOGY " & " '"
                End If
                If txt_modcode.Text = "XX" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR OTHERS " & " '"
                End If
                CrystalReport1.Formulas(7) = "PROV = '" & Trim(txt_indctinvstg.Text) & " '"
                ' showing icu in report
                CrystalReport1.Formulas(8) = "icu = '" & Trim(MCICU) & " '"
                               
                CrystalReport1.Formulas(18) = "ENTERBY='" & Trim(enter_by1) & "'"
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            End If
            '*** MICROBIOLOGY
            '*** ENDOSCOPY
             If txt_modcode = "ED" Then
                'CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPED.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & cmb_reqno.Text & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_sex & "/" & txt_Age & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PROV = '" & txt_indctinvstg & "'"
                CrystalReport1.Formulas(7) = "ENTERBY='" & Trim(enter_by1) & "'"
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
             End If
            '*** ENDOSCOPY
            '*** CARDIOLOGY
            If txt_modcode = "EK" Or txt_modcode = "2D" Or txt_modcode = "ST" Or txt_modcode = "PF" Then
                'CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPEC.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & cmb_reqno.Text & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                'CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                
                
'                If Trim(prepath) = "" Then
'                    CrystalReport1.Formulas(5) = "PATH = '" & "" & " '"
'                Else
'                    CrystalReport1.Formulas(5) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
'                End If
                If txt_modcode.Text = "EK" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR ECG" & " '"
                End If
                If txt_modcode.Text = "2D" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR 2DECHO" & " '"
                End If
                If txt_modcode.Text = "ST" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR STRESS TEST " & " '"
                End If
                If txt_modcode.Text = "PF" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR PFT TEST " & " '"
                End If
                    CrystalReport1.Formulas(7) = "ENTERBY='" & Trim(enter_by1) & "'"
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            End If
            
            
            '*** CARDIOLOGY
            '*** HAEMATOLOGY
            If txt_modcode = "HL" Then
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPHLN.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & cmb_reqno.Text & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                               
'                If Trim(prepath) = "" Then
'                    CrystalReport1.Formulas(5) = "PATH = '" & "" & " '"
'                Else
'                    CrystalReport1.Formulas(5) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
'                End If
                If txt_modcode.Text = "HL" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR HAEMATOLOGY" & " '"
                End If
                CrystalReport1.Formulas(7) = "ENTERBY='" & Trim(enter_by1) & "'"
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            End If
            
                       
            '*** HAEMATOLOGY
            '*** bone marrow aspiration
            If txt_modcode = "BM" Then
                If SELDBMACIND_FLAG = False Then
                    SELDBMACIND_QR(0) = req_chglen(cmb_reqno.Text)
                    Set SELDBMACIND_RS = SELDBMACIND_QR.OpenResultset(1, 3)
                    SELDBMACIND_FLAG = True
                Else
                    SELDBMACIND_QR(0) = req_chglen(cmb_reqno.Text)
                    SELDBMACIND_RS.MoreResults
                    SELDBMACIND_RS.Requery
                End If
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                'CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPBMN.RPT"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPHL.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & cmb_reqno.Text & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                
                If Not SELDBMACIND_RS.EOF Then
                    If SELDBMACIND_RS("NEWCASE") <> "" Then
                        CrystalReport1.Formulas(7) = "NEWCASE = '" & SELDBMACIND_RS("NEWCASE") & "'"
                    End If
                    If SELDBMACIND_RS("POSTTRANS") <> "" Then
                        CrystalReport1.Formulas(8) = "POSTTRANS = '" & SELDBMACIND_RS("POSTTRANS") & "'"
                    End If
                    If SELDBMACIND_RS("DAY10MARR") <> "" Then
                        CrystalReport1.Formulas(9) = "DAY10MARR = '" & SELDBMACIND_RS("DAY10MARR") & "'"
                    End If
                    If SELDBMACIND_RS("PARTREATED") <> "" Then
                        CrystalReport1.Formulas(10) = "PARTREATED = '" & SELDBMACIND_RS("PARTREATED") & "'"
                    End If
                    If SELDBMACIND_RS("TSTEROIDS") <> "" Then
                        CrystalReport1.Formulas(11) = "TSTEROIDS = '" & SELDBMACIND_RS("TSTEROIDS") & "'"
                    End If
                    If SELDBMACIND_RS("RELAPSE") <> "" Then
                        CrystalReport1.Formulas(12) = "RELAPSE = '" & SELDBMACIND_RS("RELAPSE") & "'"
                    End If
                    If SELDBMACIND_RS("PINDUCTION") <> "" Then
                        CrystalReport1.Formulas(13) = "PINDUCTION = '" & SELDBMACIND_RS("PINDUCTION") & "'"
                    End If
                    If SELDBMACIND_RS("POSTBMT") <> "" Then
                        CrystalReport1.Formulas(14) = "POSTBMT = '" & SELDBMACIND_RS("POSTBMT") & "'"
                    End If
                    If SELDBMACIND_RS("OTHER") <> "" Then
                        CrystalReport1.Formulas(15) = "OTHER = 'Other : " & SELDBMACIND_RS("OTHER") & "'"
                    End If
                End If
                CrystalReport1.Formulas(16) = "PROV = '" & txt_indctinvstg & "'"
'                If Trim(prepath) = "" Then
'                    CrystalReport1.Formulas(5) = "PATH = '" & "" & " '"
'                Else
'                    CrystalReport1.Formulas(5) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
'                End If
                If txt_modcode.Text = "BM" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR BONE MARROW ASPIRATION " & " '"
                End If
                CrystalReport1.Formulas(17) = "ENTERBY='" & Trim(enter_by1) & "'"
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            End If
            '*** bone marrow aspiration
            '*** BODY FLUIDS
            If txt_modcode = "BF" Then
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPBFN.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & cmb_reqno.Text & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                
               If txt_modcode.Text = "BF" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR BODY FLUIDS " & " '"
                End If
                CrystalReport1.Formulas(7) = "ENTERBY='" & Trim(enter_by1) & "'"
                
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            End If
            ''**** BLOOD PRODUCT
            If txt_modcode = "BP" Then
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPBPRP1.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & cmb_reqno.Text & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
              '  CRYSTALREPORT1.Formulas(5) = "PATHNO = '" & txt_pathno.Text & "'"
                
                
'                If Trim(prepath) = "" Then
'                    CrystalReport1.Formulas(5) = "PATH = '" & "" & " '"
'                Else
'                    CrystalReport1.Formulas(5) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
'                End If
               ' CRYSTALREPORT1.Formulas(6) = "CAP = '" & "REQUISITION FOR BLOOD PRODUCTS" & " '"
                CrystalReport1.Formulas(7) = "usrid = '" & str & " '"
                If Not SELDTMNO_RS.EOF Then
                    If IsNull(SELDTMNO_RS("DTMNO")) Then
                        CrystalReport1.Formulas(8) = "DTMNO = '" & " " & " '"
                    Else
                        CrystalReport1.Formulas(8) = "DTMNO = '" & Trim(SELDTMNO_RS("DTMNO")) & " '"
                    End If
                    CrystalReport1.Formulas(9) = "ENTERBY = '" & Trim(enter_by1) & "'"
                End If
                CrystalReport1.Formulas(10) = "PROV = '" & txt_indctinvstg & "'"
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            End If
            
            
            '*** BODY FLUIDS
            '*** HLA TYPING
            If txt_modcode = "HT" Then
                If SELDREQUISITNFLAG = False Then
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                    SELDREQUISITNFLAG = False
                Else
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    seldrequisitn_rs.MoreResults
                    seldrequisitn_rs.Requery
                End If
                If SELTA101FALG = False Then
                    'selta101
                    SELTA101_QR(0) = Msk_caseno.FormattedText
                    Set SELTA101_RS = SELTA101_QR.OpenResultset(2, 3)
                    SELTA101FALG = True
                Else
                    'selta101
                    SELTA101_QR(0) = Msk_caseno.FormattedText
                    Set SELTA101_RS = SELTA101_QR.OpenResultset(2, 3)
                    SELTA101_RS.MoreResults
                    SELTA101_RS.Requery
                End If
                If Trim(seldrequisitn_rs("LABNO")) = "yes" Then
                    CrystalReport1.Formulas(6) = "NAME = '(Donor)   " & Trim(seldrequisitn_rs("NMETERIALS")) & "'"
                    CrystalReport1.Formulas(7) = "RELATION = 'Relation to patient' "
                    CrystalReport1.Formulas(3) = "SEXAGE = '" & Trim(seldrequisitn_rs("OUTSSPECIM")) & "/" & Trim(seldrequisitn_rs("LATERALITM")) & "'"
                    CrystalReport1.Formulas(8) = "PNAME = '(Patient) " & Trim(SELTA101_RS("FIRST_NAME")) + " " + Trim(SELTA101_RS("MID_NAME")) + " " + Trim(SELTA101_RS("LAST_NAME")) & "'"
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPHTDON.RPT"
                
                Else
                    CrystalReport1.Formulas(6) = "NAME = '" & Trim(SELTA101_RS("FIRST_NAME")) + " " + Trim(SELTA101_RS("MID_NAME")) + " " + Trim(SELTA101_RS("LAST_NAME")) & "'"
                    CrystalReport1.Formulas(7) = "RELATION ='" & " " & "'"
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPHT.RPT"
                    CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                End If

                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                CrystalReport1.Formulas(0) = "REQNO = '" & cmb_reqno.Text & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                'If Len(seldrequisitn_rs("LABNO")) > 0 Then
                'If Not SELDTMNO_RS.EOF Then
'                    If IsNull(SELDTMNO_RS("DTMNO")) Then
'                        CrystalReport1.Formulas(8) = "DTMNO = '" & " " & " '"
'                    Else
'                        CrystalReport1.Formulas(8) = "DTMNO = '" & Trim(SELDTMNO_RS("DTMNO")) & " '"
'                    End If
'                End If
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            End If
            '*** HLA TYPING
            '*** DTM LAB
            
            '*** DTM LAB
        Else
            Data1.Recordset.MoveFirst
            While Not Data1.Recordset.EOF
                If SELRISSTAT_FLAG = False Then
                     SELRISSTAT_QR(0) = req_chglen(Data1.Recordset.Fields("REQNO"))
                     Set SELRISSTAT_RS = SELRISSTAT_QR.OpenResultset(1, 3)
                     SELRISSTAT_FLAG = True
                Else
                     SELRISSTAT_QR(0) = req_chglen(Data1.Recordset.Fields("REQNO"))
                     SELRISSTAT_RS.MoreResults
                     SELRISSTAT_RS.Requery
                End If
                If Not SELRISSTAT_RS.EOF Then
              '1
                    If Trim(txt_modcode.Text) = "BP" Then   '1
                        Dim X
                        X = ""
                        If SELDMULPROCEFLAG = False Then
                             SELDMULPROCEQR(0) = req_chglen(Data1.Recordset.Fields("REQNO"))
                             Set SELDMULPROCERS = SELDMULPROCEQR.OpenResultset(1, 3)
                             SELDMULPROCEFLAG = True
                        Else
                             SELDMULPROCEQR(0) = req_chglen(Data1.Recordset.Fields("REQNO"))
                             SELDMULPROCERS.MoreResults
                             SELDMULPROCERS.Requery
                         End If
                         If Not SELDMULPROCERS.EOF Then     '2
                            Do While Not SELDMULPROCERS.EOF
                                If SELDMULPROCERS("PRO_CODE") = "H002" Then
                                    X = "YES"
                                End If
                                SELDMULPROCERS.MoveNext
                            Loop
                         End If                             '2
                        If (Trim(txt_catgory.Text) = "NC" Or Trim(txt_catgory.Text) = "SF") Then  '3
                         
                         Else
                         ' THIS LINE WAS CHANGED BECAUSE IN C CASES THE SYSTEM ASKING TO PREPARE MEMO EVEN AFTER
                         ' THE MEMO EVEN AFTER THE MENO IS PREPARED
                         'If Trim(X) = "YES" And (IsNull(SELRISSTAT_RS("memono")) Or txt_catgory.Text <> "NC") Then
                         If Trim(X) = "YES" And (IsNull(SELRISSTAT_RS("memono"))) Then
                            MsgBox " Please Prepare Memo !!!"
                            Exit Sub
                         End If
                        End If   '3
                    End If          '2
                    If Trim(txt_modcode.Text) <> "BP" Then    '4
                        If Trim(txt_catgory.Text) = "NC" Or opt_staff.Value = True Then   '5
                        Else
                            If IsNull(SELRISSTAT_RS("memono")) And Trim(SELRISSTAT_RS("REQNSTAT")) <> "P" Then
                                MsgBox " Please Prepare Memo !!!"
                                Exit Sub
                            End If
                        End If      '5
                    End If      '4
                End If          '1
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(Data1.Recordset.Fields("REQNO")) & "'"
                '**** SP REPORT
                If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "SP" Then
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPSP.RPT"
                    CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
                    I = InStr(cmb_modality.Text, "$")
                    CrystalReport1.Formulas(1) = "MOD = '" & Left(cmb_modality.Text, I - 1) & "'"
                    If txt_status = "OP" Then
                        CrystalReport1.Formulas(1) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                    Else
                        CrystalReport1.Formulas(1) = "ST = '" & txt_catgory & "/" & " In Patient" & "'"
                    End If
                    '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                    '*****
                    CrystalReport1.Formulas(2) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                    If Trim(txt_status.Text) = "IP" Then
                        CrystalReport1.Formulas(3) = "WRBED = '" & txt_sex & "/" & txt_Age & "'"
                    Else
                        CrystalReport1.Formulas(3) = "WRBED = '" & " " & "'"
                    End If
                    CrystalReport1.Formulas(4) = "USRID = '" & str & "'"
                    CrystalReport1.Formulas(5) = "PROV = '" & txt_indctinvstg & "'"
                    CrystalReport1.Formulas(6) = "PATHNO = '" & txt_sppathno & "'"
                    I = InStr(cmb_sites, "$")
                    If I = 0 Then
                     If Len(Trim(cmb_sites)) > 0 Then
                        CrystalReport1.Formulas(7) = "SITE = '" & UCase(Trim(cmb_sites.Text)) & "'"
                     Else
                        CrystalReport1.Formulas(7) = "SITE = '" & "" & "'"
                     End If
                    Else
                     CrystalReport1.Formulas(7) = "SITE = '" & Mid(cmb_sites, 1, I - 1) & "'"
                    End If
                    '*****
                    If SELDREQRS.RowCount > 0 Then
                        prepath = ""
                        SELDREQRS.MoveFirst
                        While Not SELDREQRS.EOF
                                If Mid(SELDREQRS("reqno"), 4, 2) = "SP" Then
                                   If Trim(SELDREQRS("PATHNO")) <> Trim(txt_sppathno.Text) Then
                                    prepath = prepath & SELDREQRS("PATHNO") & " "
                                   End If
                                End If
                            SELDREQRS.MoveNext
                        Wend
                        prepath = Trim(prepath)
                        prepath = Left(prepath, 100)
                    End If
                    
                    '*****
                    If Trim(prepath) = "" Then
                        CrystalReport1.Formulas(8) = "PATH = '" & "" & "'"
                    Else
                        CrystalReport1.Formulas(8) = "PATH = '" & "Previous Path No." & Left(Trim(prepath), 200) & "'"
                    End If
                    CrystalReport1.Formulas(9) = "ENTERBY='" & Trim(enter_by1) & "'"
               End If
               '**** SS REPORT
               
             If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "SS" Then
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(Data1.Recordset.Fields("REQNO")) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPSS.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & TEXT1.Text & "'" ' txt_pathno & "'"


'                If Trim(prepath) = "" Then
'                    CrystalReport1.Formulas(5) = "PATH = '" & "" & " '"
'                Else
'                    CrystalReport1.Formulas(5) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
'                End If
                If txt_modcode.Text = "SS" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR SCREENING SERVICES" & " '"
                End If
                If Not SELDTMNO_RS.EOF Then
                    If IsNull(SELDTMNO_RS("DTMNO")) Then
                        CrystalReport1.Formulas(8) = "DTMNO = '" & " " & " '"
                    Else
                        CrystalReport1.Formulas(8) = "DTMNO = '" & Trim(SELDTMNO_RS("DTMNO")) & " '"
                    End If
                End If
'                CRYSTALREPORT1.Action = 1
'                Exit Sub
            End If
                '**** SP REPORT
                 '**** IH REPORT
                 If Mid(Data1.Recordset.Fields("REQNO"), 5, 2) = "IH" Then
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPIHN.RPT"
                    CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
                    I = InStr(cmb_modality.Text, "$")
                    CrystalReport1.Formulas(1) = "MOD = '" & Left(cmb_modality.Text, I - 1) & "'"
                    If txt_status = "OP" Then
                        CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                    Else
                        CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & "In Patient" & "'"
                    End If
                    '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                    '*****
                    CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                    If Trim(txt_status.Text) = "IP" Then
                        CrystalReport1.Formulas(4) = "WRBED = '" & txt_sex & "/" & txt_Age & "'"
                    Else
                        CrystalReport1.Formulas(5) = "WRBED = '" & " " & "'"
                    End If
                     CrystalReport1.Formulas(6) = "outblock = '" & cmb_Ihsoutsideblock & "'"
                     CrystalReport1.Formulas(7) = "outslide = '" & cmb_Ihsoutslide & "'"

                     If seldihreportflag = False Then
                    seldihreport_qr(0) = req_chglen(txt_ihreqno)
                    Set seldihreport_rs = seldihreport_qr.OpenResultset(1, 3)
                    seldihreportflag = True
                 Else
                    seldihreport_qr(0) = req_chglen(cmb_reqno.Text)
                    seldihreport_rs.MoreResults
                    seldihreport_rs.Requery
                 End If
                 If seldihreport_rs.RowCount > 0 Then
                 str1 = ""
                    While Not seldihreport_rs.EOF
                        str1 = str1 & Trim(seldihreport_rs("testorder")) & ", "
                        seldihreport_rs.MoveNext
                    Wend
                 End If
                    I = Len(Trim(str1))
                    If I = 0 Then
                        CrystalReport1.Formulas(8) = "TEST = '" & "" & "'"
                    Else
                        CrystalReport1.Formulas(8) = "TEST = '" & Left(Trim(str1), I - 1) & "'"
                    End If
                    CrystalReport1.Formulas(9) = "USRID = '" & str & "'"
                    CrystalReport1.Formulas(10) = "PROV = '" & txt_indctinvstg & "'"
                    CrystalReport1.Formulas(11) = "PATHNO = '" & txt_ihpathno & "'"
                    If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(12) = "PATH = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(12) = "PATH = '" & "Previous Path No." & Left(Trim(prepath), 200) & "'"
                End If
                    CrystalReport1.Formulas(13) = "ENTERBY='" & Trim(enter_by1) & "'"
                 End If
                 
                 '******HLA REPORT
                 If Mid(Data1.Recordset.Fields("REQNO"), 5, 2) = "HT" Then
                     'If txt_modcode = "HT" Then
                    
                    'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(Data1.Recordset.Fields("REQNO")) & "'"
                    CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
                    I = InStr(cmb_modality.Text, "$")
                    CrystalReport1.Formulas(1) = "MOD = '" & Left(cmb_modality.Text, I - 1) & "'"
                    If SELDREQUISITNFLAG = False Then
                        seldrequisitn_qr(0) = req_chglen(Data1.Recordset.Fields("REQNO"))
                        Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                        SELDREQUISITNFLAG = False
                    Else
                        seldrequisitn_qr(0) = req_chglen(Data1.Recordset.Fields("REQNO"))
                        seldrequisitn_rs.MoreResults
                        seldrequisitn_rs.Requery
                    End If
                    If SELTA101FALG = False Then
                        'selta101
                        SELTA101_QR(0) = Msk_caseno.FormattedText
                        Set SELTA101_RS = SELTA101_QR.OpenResultset(2, 3)
                        SELTA101FALG = True
                    Else
                        'selta101
                        SELTA101_QR(0) = Msk_caseno.FormattedText
                        Set SELTA101_RS = SELTA101_QR.OpenResultset(2, 3)
                        SELTA101_RS.MoreResults
                        SELTA101_RS.Requery
                    End If
                    If txt_status = "OP" Then
                        CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                    Else
                        CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                    End If
                    '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                    '*****
                    If Trim(txt_status.Text) = "IP" Then
                        CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                    Else
                        CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                    End If
                    CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                    'If Len(seldrequisitn_rs("LABNO")) > 0 Then
                    If Trim(seldrequisitn_rs("labno")) = "yes" Then
                        CrystalReport1.Formulas(6) = "NAME = '(Donor)    " & seldrequisitn_rs("NMETERIALS") & "'"
                        CrystalReport1.Formulas(7) = "RELATION = 'Relation to patient' "
                        CrystalReport1.Formulas(3) = "SEXAGE = '" & Trim(seldrequisitn_rs("OUTSSPECIM")) & "/" & Trim(seldrequisitn_rs("LATERALITM")) & "'"
                        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPHTdon.RPT"
                        CrystalReport1.Formulas(8) = "PNAME = '(Patient) " & Trim(SELTA101_RS("FIRST_NAME")) + " " + Trim(SELTA101_RS("MID_NAME")) + " " + Trim(SELTA101_RS("LAST_NAME")) & "'"
                    Else
                        CrystalReport1.Formulas(6) = "NAME = '" & Trim(SELTA101_RS("FIRST_NAME")) + " " + Trim(SELTA101_RS("MID_NAME")) + " " + Trim(SELTA101_RS("LAST_NAME")) & "'"
                        CrystalReport1.Formulas(7) = "RELATION ='" & " " & "'"
                        CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPHT.RPT"
                    End If
'                    If Not SELDTMNO_RS.EOF Then
'                       If IsNull(SELDTMNO_RS("DTMNO")) Then
'                           CrystalReport1.Formulas(8) = "DTMNO = '" & " " & " '"
'                       Else
'                           CrystalReport1.Formulas(8) = "DTMNO = '" & Trim(SELDTMNO_RS("DTMNO")) & " '"
'                       End If
'                   End If
                    'CRYSTALREPORT1.Action = 1
                    'Exit Sub
                End If
                
                 '**** IH REPORT
                 '**** FA REPORT
                 If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "FA" Then
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPFAN.RPT"
                    CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
'                    I = InStr(cmb_modality.Text, "$")
'                    CRYSTALREPORT1.Formulas(1) = "MOD = '" & Left(cmb_modality.Text, I - 1) & "'"
                    If txt_status = "OP" Then
                        CrystalReport1.Formulas(1) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                    Else
                        CrystalReport1.Formulas(1) = "ST = '" & txt_catgory & "/" & " In Patient" & "'"
                    End If
                    '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                    '*****
                    CrystalReport1.Formulas(2) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                    If Trim(txt_status.Text) = "IP" Then
                        CrystalReport1.Formulas(3) = "WRBED = '" & txt_sex & "/" & txt_Age & "'"
                    Else
                        CrystalReport1.Formulas(3) = "WRBED = '" & " " & "'"
                    End If
                    CrystalReport1.Formulas(4) = "USRID = '" & str & "'"
                    CrystalReport1.Formulas(5) = "PROV = '" & txt_indctinvstg & "'"
                    CrystalReport1.Formulas(6) = "PATHNO = '" & txt_sppathno & "'"
                    I = InStr(cmb_fnsite, "$")
                    If I = 0 Then
                     If Len(Trim(cmb_fnsite)) > 0 Then
                        CrystalReport1.Formulas(7) = "SITE = '" & UCase(Trim(cmb_fnsite.Text)) & "'"
                     Else
                        CrystalReport1.Formulas(7) = "SITE = '" & "" & "'"
                     End If
                    Else
                     CrystalReport1.Formulas(7) = "SITE = '" & Mid(cmb_fnsite, 1, I - 1) & "'"
                    End If
                    If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(8) = "PATH = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(8) = "PATH = '" & "Previous Path No." & Left(Trim(prepath), 200) & "'"
                End If
                    CrystalReport1.Formulas(19) = "ENTERBY='" & Trim(enter_by1) & "'"
               End If
                 '**** FA REPORT
                 '**** CG OR CN REPORT
                 If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "CG" Or Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "CN" Or Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "PG" Or Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "PN" Then
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPCGN.RPT"
                    CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
'                    I = InStr(cmb_modality.Text, "$")
'                    CRYSTALREPORT1.Formulas(1) = "MOD = '" & Left(cmb_modality.Text, I - 1) & "'"
                    If txt_status = "OP" Then
                        CrystalReport1.Formulas(1) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                    Else
                        CrystalReport1.Formulas(1) = "ST = '" & txt_catgory & "/" & " In Patient" & "'"
                    End If
                    '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                    '*****
                    CrystalReport1.Formulas(2) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                    If Trim(txt_status.Text) = "IP" Then
                        CrystalReport1.Formulas(3) = "WRBED = '" & txt_sex & "/" & txt_Age & "'"
                    Else
                        CrystalReport1.Formulas(3) = "WRBED = '" & " " & "'"
                    End If
                    CrystalReport1.Formulas(4) = "USRID = '" & str & "'"
                    CrystalReport1.Formulas(5) = "PROV = '" & txt_indctinvstg & "'"
                    CrystalReport1.Formulas(6) = "PATHNO = '" & txt_sppathno & "'"
                    I = InStr(cmb_cnsite, "$")
                    If I = 0 Then
                     If Len(Trim(cmb_cnsite)) > 0 Then
                        CrystalReport1.Formulas(7) = "SITE = '" & UCase(Trim(cmb_cnsite.Text)) & "'"
                     Else
                        CrystalReport1.Formulas(7) = "SITE = '" & "" & "'"
                     End If
                    Else
                     CrystalReport1.Formulas(7) = "SITE = '" & Mid(cmb_cnsite, 1, I - 1) & "'"
                    End If
                    If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(8) = "PATH = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(8) = "PATH = '" & "Previous Path No." & Left(Trim(prepath), 200) & "'"
                End If
                    CrystalReport1.Formulas(9) = "ENTERBY='" & Trim(enter_by1) & "'"
               End If
               
                 '**** CG OR CN REPORT
                 '**** BC OR EL OR TM REPORT
                If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "BC" Or Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "EL" Or Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "TM" Then
                'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(Data1.Recordset.Fields("REQNO")) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPBC.RPT"
                'CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
                CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
''                    i = InStr(cmb_modality.Text, "$")
''                    CrystalReport1.Formulas(1) = "MOD = '" & Left(cmb_modality.Text, i - 1) & "'"
                    If txt_status = "OP" Then
                        CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                    Else
                        CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & "In Patient" & "'"
                    End If
                    '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                    '*****
                    CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                    If Trim(txt_status.Text) = "IP" Then
                        CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                    Else
                        CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                    End If
                '********** Added by Shraddha Pandit on 08/01/09 ************
                CrystalReport1.Formulas(9) = "PROV = '" & txt_clinhist & "'"
                '************************************************************
'                If Trim(prepath) = "" Then
'                    CrystalReport1.Formulas(6) = "PATH = '" & "" & " '"
'                Else
'                    CrystalReport1.Formulas(6) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
'                End If
                If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "BC" Then
                   CrystalReport1.Formulas(7) = "CAP = '" & "REQUISITION FOR BIOCHEMSITRY" & " '"
                End If
                If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "EL" Then
                   CrystalReport1.Formulas(7) = "CAP = '" & "REQUISITION FOR EMERGENCY LAB" & " '"
                End If
                If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "TM" Then
                   CrystalReport1.Formulas(7) = "CAP = '" & "REQUISITION FOR TUMOUR MARKER " & " '"
                End If
                CrystalReport1.Formulas(8) = "ENTERBY = '" & Trim(enter_by1) & "'"
            End If
                 '**** BC OR EL OR TM REPORT
                 '**** MICROBIOLOGY
                 If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "BA" Or Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "MB" Or Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "SE" Or Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "MC" Or Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "CM" Or Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "MO" Or Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "XX" Then
                
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPMB.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
               ' For k1 = 0 To cmb_modality.ListCount
                    I = InStr(cmb_modality.Text, "$")
                '    If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = Trim(Mid(cmb_modality.List(k1), i + 1)) Then
                 '         CRYSTALREPORT1.Formulas(1) = "MOD = '" & Left(cmb_modality.Text, i - 1) & "'"
                 '   End If
                'Next k1
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & Trim(txt_ward) & "/" & Trim(txt_bedno.Text) & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                
                
                If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(5) = "PATH = '" & "" & " '"
                Else
                    CrystalReport1.Formulas(5) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
                End If
                If txt_modcode.Text = "BA" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR BACTERIOLOGY" & " '"
                End If
                If txt_modcode.Text = "MB" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR MYCOBACTERIOLOGY" & " '"
                End If
                If txt_modcode.Text = "SE" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR SEROLOGY " & " '"
                End If
                If txt_modcode.Text = "MC" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR MYCOLOGY " & " '"
                End If
                If txt_modcode.Text = "CM" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR CLINICAL MICROBIOLOGY" & " '"
                End If
                If txt_modcode.Text = "MO" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR MOLECULAR MICROBIOLOGY " & " '"
                End If
                If txt_modcode.Text = "XX" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR OTHERS " & " '"
                End If
                CrystalReport1.Formulas(7) = "PROV = '" & Trim(txt_indctinvstg.Text) & "'"
                CrystalReport1.Formulas(8) = "icu = '" & Trim(MCICU) & " '"
                CrystalReport1.Formulas(18) = "ENTERBY = '" & Trim(enter_by1) & "'"
            End If
                 '**** MICROBIOLOGY
                 '**** ENDOSCOPY
             If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "ED" Then
                    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(Data1.Recordset.Fields("REQNO")) & "'"
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPEDN.RPT"
                    CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
                    If txt_status = "OP" Then
                        CrystalReport1.Formulas(1) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                    Else
                        CrystalReport1.Formulas(1) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                    End If
                    '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                    '*****
                    CrystalReport1.Formulas(2) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                    If Trim(txt_status.Text) = "IP" Then
                        CrystalReport1.Formulas(3) = "WRBED = '" & txt_sex & "/" & txt_Age & "'"
                    Else
                        CrystalReport1.Formulas(3) = "WRBED = '" & " " & "'"
                    End If
                    CrystalReport1.Formulas(4) = "PROV = '" & txt_indctinvstg & "'"
                    CrystalReport1.Formulas(18) = "ENTERBY = '" & Trim(enter_by1) & "'"
              End If
                 
                 '**** ENDOSCOPY
                 '**** CARDIOLOGY
                 If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "EK" Or Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "2D" Or Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "ST" Or Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "PF" Then
                    'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(Data1.Recordset.Fields("REQNO")) & "'"
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPECN.RPT"
                    'CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
                    CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
    ''                    i = InStr(cmb_modality.Text, "$")
    ''                    CrystalReport1.Formulas(1) = "MOD = '" & Left(cmb_modality.Text, i - 1) & "'"
                        If txt_status = "OP" Then
                           CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                        Else
                            CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & "In Patient" & "'"
                        End If
                        '********  on 31/12/04
                        If txt_status = "ST" Then
                            CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                        End If
                        '*****
                        CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                        If Trim(txt_status.Text) = "IP" Then
                            CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                        Else
                            CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                        End If
                    
    '                If Trim(prepath) = "" Then
    '                    CrystalReport1.Formulas(6) = "PATH = '" & "" & " '"
    '                Else
    '                    CrystalReport1.Formulas(6) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
    '                End If
                    If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "EK" Then
                       CrystalReport1.Formulas(7) = "CAP = '" & "REQUISITION FOR ECG" & " '"
                    End If
                    If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "2D" Then
                       CrystalReport1.Formulas(7) = "CAP = '" & "REQUISITION FOR 2DECHO" & " '"
                    End If
                    If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "ST" Then
                       CrystalReport1.Formulas(7) = "CAP = '" & "REQUISITION FOR STRESS TEST " & " '"
                    End If
                    If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "PF" Then
                       CrystalReport1.Formulas(7) = "CAP = '" & "REQUISITION FOR PFT  " & " '"
                    End If
                    CrystalReport1.Formulas(9) = "ENTERBY = '" & Trim(enter_by1) & "'"
            End If
                 
                 '**** CARDIOLOGY
                 '**** HAEMATOLOGY
                 If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "HL" Then
                'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPHLN.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                    '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                
                
'                If Trim(prepath) = "" Then
'                    CrystalReport1.Formulas(5) = "PATH = '" & "" & " '"
'                Else
'                    CrystalReport1.Formulas(5) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
'                End If
                If txt_modcode.Text = "HL" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR HAEMATOLOGY" & " '"
                End If
                CrystalReport1.Formulas(18) = "ENTERBY = '" & Trim(enter_by1) & "'"
            End If
                 '**** HAEMATOLOGY
                 '**** BONE MARROW ASPIRATION
            If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "BM" Then
                If SELDBMACIND_FLAG = False Then
                    SELDBMACIND_QR(0) = req_chglen(Data1.Recordset.Fields("REQNO"))
                    Set SELDBMACIND_RS = SELDBMACIND_QR.OpenResultset(1, 3)
                    SELDBMACIND_FLAG = True
                Else
                    SELDBMACIND_QR(0) = req_chglen(Data1.Recordset.Fields("REQNO"))
                    SELDBMACIND_RS.MoreResults
                    SELDBMACIND_RS.Requery
                End If
                'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                'CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPHLN.RPT"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPHL.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                
                If Not SELDBMACIND_RS.EOF Then
                    If SELDBMACIND_RS("NEWCASE") <> "" Then
                        CrystalReport1.Formulas(7) = "NEWCASE = '" & SELDBMACIND_RS("NEWCASE") & "'"
                    End If
                    If SELDBMACIND_RS("POSTTRANS") <> "" Then
                        CrystalReport1.Formulas(8) = "POSTTRANS = '" & SELDBMACIND_RS("POSTTRANS") & "'"
                    End If
                    If SELDBMACIND_RS("DAY10MARR") <> "" Then
                        CrystalReport1.Formulas(9) = "DAY10MARR = '" & SELDBMACIND_RS("DAY10MARR") & "'"
                    End If
                    If SELDBMACIND_RS("PARTREATED") <> "" Then
                        CrystalReport1.Formulas(10) = "PARTREATED = '" & SELDBMACIND_RS("PARTREATED") & "'"
                    End If
                    If SELDBMACIND_RS("TSTEROIDS") <> "" Then
                        CrystalReport1.Formulas(11) = "TSTEROIDS = '" & SELDBMACIND_RS("TSTEROIDS") & "'"
                    End If
                    If SELDBMACIND_RS("RELAPSE") <> "" Then
                        CrystalReport1.Formulas(12) = "RELAPSE = '" & SELDBMACIND_RS("RELAPSE") & "'"
                    End If
                    If SELDBMACIND_RS("PINDUCTION") <> "" Then
                        CrystalReport1.Formulas(13) = "PINDUCTION = '" & SELDBMACIND_RS("PINDUCTION") & "'"
                    End If
                    If SELDBMACIND_RS("POSTBMT") <> "" Then
                        CrystalReport1.Formulas(14) = "POSTBMT = '" & SELDBMACIND_RS("POSTBMT") & "'"
                    End If
                    If SELDBMACIND_RS("OTHER") <> "" Then
                        CrystalReport1.Formulas(15) = "OTHER = 'Other : " & SELDBMACIND_RS("OTHER") & "'"
                    End If
                End If
                CrystalReport1.Formulas(16) = "PROV = '" & txt_indctinvstg & "'"
'                If Trim(prepath) = "" Then
'                    CrystalReport1.Formulas(5) = "PATH = '" & "" & " '"
'                Else
'                    CrystalReport1.Formulas(5) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
'                End If
                If txt_modcode.Text = "BM" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR BONE MARROW ASPIRATION" & " '"
                End If
                CrystalReport1.Formulas(18) = "ENTERBY='" & Trim(enter_by1) & "'"
              End If
               ''******BLOOD PRODUCT
               'HERE
              If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "BP" Then
'                If AMFLAG = False Then
'                Ii = Data1.Recordset.Fields("REQNO")
'                   AMQR(0) = Ii
'                   Set AMRS = AMQR.OpenResultset(1, 3)
'                    AMFLAG = True
'                Else
'                    AMQR(0) = req_chglen(cmb_reqno.Text)
'                    AMRS.MoreResults
'                    AMRS.Requery
'                End If
'                MsgBox STR
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(Data1.Recordset.Fields("REQNO")) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPBPRP1.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                ' CODE
'             If ind4 <> "" Then
'                CRYSTALREPORT1.Formulas(5) = "QTY1='" & ind4 & "'"
'             ElseIf ind5 <> "" Then
'                CRYSTALREPORT1.Formulas(5) = "QTY1='" & ind5 & "'"
'             End If

            
            
            '    CRYSTALREPORT1.Formulas(5) = "PATHNO = '" & txt_pathno.Text & "'"
              
                
'                If Trim(prepath) = "" Then
'                    CrystalReport1.Formulas(5) = "PATH = '" & "" & " '"
'                Else
'                    CrystalReport1.Formulas(5) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
'                End If
                'If txt_modcode.Text = "SS" Then
                 '  CRYSTALREPORT1.Formulas(6) = "CAP = '" & "REQUISITION FOR BLOOD PRODUCTS" & " '"
                'End If
                '***
                CrystalReport1.Formulas(10) = "PROV = '" & txt_indctinvstg & "'"
                 CrystalReport1.Formulas(7) = "usrid = '" & str & "'"
                 If Not SELDTMNO_RS.EOF Then
                    If IsNull(SELDTMNO_RS("DTMNO")) Then
                        CrystalReport1.Formulas(8) = "DTMNO = '" & " " & "'"
                    Else
                        CrystalReport1.Formulas(8) = "DTMNO = '" & Trim(SELDTMNO_RS("DTMNO")) & "'"
                    End If
                End If
              '***
                   CrystalReport1.Formulas(9) = "ENTERBY = '" & Trim(enter_by1) & "'"
            End If
                 
                 '**** BONE MARROW ASPIRATION
                 '**** BODY FLUIDS
                If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "BF" Then
                'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPHLN.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                
'                If Trim(prepath) = "" Then
'                    CrystalReport1.Formulas(5) = "PATH = '" & "" & " '"
'                Else
'                    CrystalReport1.Formulas(5) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
'                End If
                If txt_modcode.Text = "BF" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR BODY FLUIDS" & " '"
                End If
                    CrystalReport1.Formulas(7) = "ENTERBY='" & Trim(enter_by1) & "'"
               End If
                
                 '**** BODY FLUIDS
                 ' FOR CYTPGENITICS 13-04-2005
               If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "FI" Or Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "CC" Then
                
                 CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
                
                'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                  CYPRINT
               End If
               
               If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "MP" Then
                  CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
                  'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                  MPPRINT
               End If
               
               
                ' FOR CYTOGENITICS 13-04-2005
                'CRYSTALREPORT1.WindowState = crptMaximized
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
                If Trim(txt_modcode.Text) = "BP" Then
                    MsgBox ("Please Do the sample Collection in the Collection Centre"), vbCritical
                    ''Exit Sub
                End If
                Data1.Recordset.MoveNext
                For I = 0 To 25
                    CrystalReport1.Formulas(I) = ""
                Next I
            Wend
        End If
End If
 ''' FOR CYOTOGENITICS 13-04-2005
         If viewflag = True Then
                If txt_modcode = "CC" Or txt_modcode = "FI" Then
                'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                CYPRINT
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
                End If
         End If

 ''' FOR CYOTOGENITICS 13-04-2005
If opt_staff.Value = True Then

    staffprint
End If

btnenabl ("PRINT")

''''Done by Kanchan for Barcode printing on 03/07/20008


If Trim(txt_dept.Text) <> "VZZ" Then
    If NEWFLAG_BCD = True Then
        Data1.Recordset.MoveFirst
        While Not Data1.Recordset.EOF
            'reqno = req_chglen(DAUTOREQNO)
            REQNO = req_chglen(Data1.Recordset.Fields("REQNO"))
            Barcode
            MYDELAY (1)
       
            Data1.Recordset.MoveNext
        Wend
    ElseIf VIEWFLAG_BCD = True Then
        REQNO = Trim(req_chglen(cmb_reqno.Text))
        Barcode
    End If
End If
''''Done by Kanchan

Exit Sub
'MSG:
'MsgBox Err.Description
'Data1.Recordset.MoveNext
End Sub

Private Sub cmd_repeat_Click()
'HIDE BY AM ON 10/06/04
'If lbl_icureq.Visible = True Then
'       MsgBox "Repeat Facility Not Available for ICU Requisition"
'       Exit Sub
'End If


'v_rep = True
I = Data1.Recordset.RecordCount
If I >= 11 Then
    MsgBox " The Maximum No. of Requisitions is Eleven" & Chr(13) & "The Maximum No. of Procedurecs in a Requisition is Eleven"
    CMB_MEMO.Enabled = True
    CMB_MEMO.SetFocus
End If

btnenabl ("REPEAT")
cmb_dept.Enabled = True
cmb_modality.Enabled = True
SPCLEAR
'RCLEAR
IHCLEAR
FACLEAR
CGCLEAR
MBCLEAR
surgicalPathology.Visible = False
fnac.Visible = False
CYTOLOGY.Visible = False
Microbiology.Visible = False
immunohistochemistry.Visible = False
REPEATFLAG = True
NEWFLAG = True
PRINTNEW = True
cmb_modality.ListIndex = -1
cmb_dept.SetFocus

End Sub

Private Sub cmd_save_Click()
Dim I As Integer
On Error Resume Next

On Error GoTo MSG:
If ipadd = "150.150.200.62" Or ipadd = "150.150.15.51" Then
    lbl_icureq.Visible = True
End If
If ipadd = "10.3.1.56" Then
    lbl_rwd.Visible = True
End If
If Trim(txt_dept.Text) = "" Then
    MsgBox "Department is Empty"
    txt_dept.SetFocus
    Exit Sub
End If
If Trim(txt_modcode.Text) = "" Then
    MsgBox "Modality Code is Empty."
    txt_modcode.SetFocus
    Exit Sub
End If

'If Trim(cmb_cnoslides.Text) = "" Then
'    MsgBox "Please Enter No. of slides"
'    cmb_cnoslides.SetFocus
'    Exit Sub
'End If
If Trim(txt_modcode.Text) = "SP" Then
    If Len(Trim(cmb_procedure.Text)) = 0 Then
        MsgBox "Please Select the Procedure", vbInformation
        cmb_procedure.SetFocus
        Exit Sub
    End If
End If
If Trim(txt_dept.Text) = "GZZ" Then
    If Len(Trim(txt_indctinvstg.Text)) = 0 Then
        MsgBox "Please Enter the Indication For Investigation", vbInformation
        txt_indctinvstg.SetFocus
        Exit Sub
    End If
    'If Len(Trim(txt_mbspecimen.Text)) = 0 Then
    If Len(Trim(cmb_specimen.Text)) = 0 Then
        MsgBox "Please Enter Specimen ", vbInformation
        'txt_mbspecimen.SetFocus
        cmb_specimen.SetFocus
        Exit Sub
    End If
    If Trim(txt_modcode.Text) <> "SE" Then
        If Len(Trim(txt_ongoinganti.Text)) = 0 Then
            MsgBox "Please Enter Ongoing Antimicrobial ", vbInformation
            txt_ongoinganti.SetFocus
            Exit Sub
        End If
    End If
End If
'If Trim(txt_modcode.Text) = "IH" Then
'    If Len(Trim(cmb_ihprocedure.Text)) = 0 Then
'        MsgBox "Please Select the Procedure", vbInformation
'        cmb_ihprocedure.SetFocus
'        Exit Sub
'    End If
'End If
If Trim(txt_modcode.Text) = "FA" Then
    If Len(Trim(CMB_FNPROCEDURE.Text)) = 0 Then
        MsgBox "Please Select the Procedure", vbInformation
        CMB_FNPROCEDURE.SetFocus
        Exit Sub
    End If
End If
If Trim(txt_modcode.Text) = "CG" Or Trim(txt_modcode.Text) = "CN" Then
    If Len(Trim(CMB_CGPROCEDURE.Text)) = 0 Then
        MsgBox "Please Select the Procedure", vbInformation
        CMB_CGPROCEDURE.SetFocus
        Exit Sub
    End If
End If

While InStr(txt_nmeterial, Chr(9)) > 0
    Dim beforst As String, afterst As String
    Dim K As Long
    K = InStr(txt_nmeterial, Chr(9))
    beforst = Mid(txt_nmeterial, 1, K - 1) & "      "
    afterst = Mid(txt_nmeterial, K + 1)
    txt_nmeterial = beforst & afterst
Wend

While InStr(txt_nmeterial, Chr(146)) > 0
    K = InStr(txt_nmeterial, Chr(146))
    beforst = Mid(txt_nmeterial, 1, K - 1) & "`"
    afterst = Mid(txt_nmeterial, K + 1)
    txt_nmeterial = beforst & afterst
Wend
'    While InStr(txt_ihmeterial, Chr(9)) > 0
'
'            K = InStr(txt_ihmeterial, Chr(9))
'            beforst = Mid(txt_ihmeterial, 1, K - 1) & "      "
'            afterst = Mid(txt_ihmeterial, K + 1)
'            txt_ihmeterial = beforst & afterst
'       Wend
'        While InStr(txt_ihmeterial, Chr(146)) > 0
'
'            K = InStr(txt_ihmeterial, Chr(146))
'            beforst = Mid(txt_ihmeterial, 1, K - 1) & "`"
'            afterst = Mid(txt_ihmeterial, K + 1)
'            txt_ihmeterial = beforst & afterst
'       Wend
While InStr(txt_scopyfind, Chr(9)) > 0
    K = InStr(txt_scopyfind, Chr(9))
    beforst = Mid(txt_scopyfind, 1, K - 1) & "      "
    afterst = Mid(txt_scopyfind, K + 1)
    txt_scopyfind = beforst & afterst
Wend

While InStr(txt_scopyfind, Chr(146)) > 0
    K = InStr(txt_scopyfind, Chr(146))
    beforst = Mid(txt_scopyfind, 1, K - 1) & "`"
    afterst = Mid(txt_scopyfind, K + 1)
    txt_scopyfind = beforst & afterst
Wend

While InStr(txt_trtdetail, Chr(9)) > 0
    K = InStr(txt_trtdetail, Chr(9))
    beforst = Mid(txt_trtdetail, 1, K - 1) & "      "
    afterst = Mid(txt_trtdetail, K + 1)
    txt_trtdetail = beforst & afterst
Wend

While InStr(txt_trtdetail, Chr(146)) > 0
    K = InStr(txt_trtdetail, Chr(146))
    beforst = Mid(txt_trtdetail, 1, K - 1) & "`"
    afterst = Mid(txt_trtdetail, K + 1)
    txt_trtdetail = beforst & afterst
Wend
       
If Trim(txt_modcode.Text) = "2D" Or Trim(txt_modcode.Text) = "PF" Or Trim(txt_modcode.Text) = "ST" Then
    If CHKFLAG = False Then
        MsgBox "You are not Selected any Indications, You Can select at least one", vbYesNo
        Exit Sub
    End If
End If

If Trim(txt_modcode.Text) = "ED" Then
    If Len(Trim(cmb_edprocedureheading.Text)) = 0 Then
        MsgBox "Procedure heading is a mandatory field."
        Exit Sub
    End If
        
    If Len(Trim(cmb_hpprocedure.Text)) = 0 Then
        MsgBox "Procedure code is empty "
        Exit Sub
    End If
        
End If

rdb.BeginTrans
If NEWFLAG = True Then
    If Trim(txt_indctinvstg.Text) = "" Then
        MsgBox " Provisional Diagnosis should not be empty"
        txt_indctinvstg.SetFocus
        Exit Sub
    End If
    
    CMD_ADDNEW_Click
    REQNOAUTO
    ' codes for cytogenitics 13-04-2005
 
    If Trim(txt_modcode.Text) = "CC" Or Trim(txt_modcode.Text) = "FI" Then
       If Trim(cmb_cyprocedure.Text) = "" Then
          MsgBox "Select procedure", vbCritical
          Exit Sub
        End If
        CYASSIGN
    End If
   
    ' codes for cytogenitics 13-04-2005

    If Trim(txt_modcode.Text) = "MP" Then
       MPASSIGN
    End If
    
    
'    If Trim(txt_modcode.Text) = "SP" Then
'        PATHNOAUTO
'    End If
'    If Trim(txt_modcode.Text) = "BA" Or Trim(txt_modcode.Text) = "MB" Or Trim(txt_modcode.Text) = "SE" Or Trim(txt_modcode.Text) = "MC" Or Trim(txt_modcode.Text) = "CM" Or Trim(txt_modcode.Text) = "MO" Or Trim(txt_modcode.Text) = "XX" Then
'        PATHNOAUTO
'    End If
    If Trim(txt_modcode.Text) = "SS" Then
        PATHNOAUTO
    End If
    If txt_modcode.Text = "SP" Then
        SPASSIGN
    End If
'    If txt_modcode.Text = "IH" Then ' commented by chandu 3/01/2003
'        IHASSIGN
'    End If
    If txt_modcode.Text = "FA" Then
        FAASSIGN
    End If
    If Trim(txt_modcode.Text) = "CG" Or Trim(txt_modcode.Text) = "CN" Then 'Or Trim(txt_modcode.Text) = "PG" Or Trim(txt_modcode.Text) = "PN" Then
        CGASSIGN
    End If
    If Trim(txt_modcode.Text) = "BC" Then
        BCASSIGN
    End If
    If Trim(txt_modcode.Text) = "EL" Then
        ELASSIGN
    End If
    If Trim(txt_modcode.Text) = "TM" Then
        TMASSIGN
    End If
    If Trim(txt_modcode.Text) = "BA" Or Trim(txt_modcode.Text) = "MB" Or Trim(txt_modcode.Text) = "SE" Or Trim(txt_modcode.Text) = "MC" Or Trim(txt_modcode.Text) = "CM" Or Trim(txt_modcode.Text) = "MO" Or Trim(txt_modcode.Text) = "XX" Then
        MBASSIGN
    End If
    If Trim(txt_modcode.Text) = "ED" Then
        EDASSIGN
    End If
    If Trim(txt_modcode.Text) = "2D" Or Trim(txt_modcode.Text) = "EK" Or Trim(txt_modcode.Text) = "PF" Or Trim(txt_modcode.Text) = "ST" Then
        ECASSIGN
    End If
    If Trim(txt_modcode.Text) = "HL" Then
        HLASSIGN
    End If
    If Trim(txt_modcode.Text) = "BM" Then
        BMASSIGN
    End If
    If Trim(txt_modcode.Text) = "BF" Then
        BFASSIGN
    End If
    If Trim(txt_modcode.Text) = "SS" Then
        If Trim(txt_indctinvstg.Text) = "" Then
           MsgBox " Provisional Diagnosis should not be empty"
           Exit Sub
        End If
        SSASSIGN
    End If
    If Trim(txt_modcode.Text) = "BP" Then
       If Trim(txt_indctinvstg.Text) = "" Then
           MsgBox " Provisional Diagnosis should not be empty"
           txt_indctinvstg.SetFocus
           Exit Sub
       End If
       BPASSIGN
    End If
    If Trim(txt_modcode.Text) = "HT" Then
        If Trim(txt_indctinvstg.Text) = "" Then
           MsgBox " Provisional Diagnosis should not be empty"
           Exit Sub
        End If
        HTASSIGN
    End If
    
'    If Trim(Len((txt_indctinvstg.Text))) = 0 Then
'        MsgBox "PLEASE ENTER THE DETAILS FOR INDICATION FOR INVESTIGATION!!", vbInformation
'        Exit Sub
'    End If
    
    'INSERT TIME AND DATE IN RTIMELOG
    INSTIME_QR(0) = req_chglen(txt_bcreqno.Text)
    INSTIME_QR(1) = Msk_caseno.FormattedText
    INSTIME_QR(2) = txt_date.Text
    INSTIME_QR(3) = txt_time.Text '
    INSTIME_QR.Execute

    '''''''DONE BY KANCHAN

    If CMB_TEXT.ListCount <> 0 Then
        For I = 0 To CMB_TEXT.ListCount - 1
            If Trim(CMB_TEXT.List(I)) = "F833" Or Trim(CMB_TEXT.List(I)) = "F834" Or Trim(CMB_TEXT.List(I)) = "F835" Or Trim(CMB_TEXT.List(I)) = "F836" Then
                
                UPDRTIMEFORBC(0) = datereq
                UPDRTIMEFORBC(1) = timereq
                UPDRTIMEFORBC(2) = req_chglen(txt_bcreqno.Text)
                UPDRTIMEFORBC.Execute
                        
            End If
        Next I
    End If


'    If ipadd = "150.150.200.62" Or ipadd = "150.150.14.171" Then
'        If Trim(CMB_ICU1.Text) <> "" Then
'
'            If Trim(CMB_ICU1.Text) = "YES" Then
'                UPDICU_QR(0) = "YES"
'            Else
'                UPDICU_QR(0) = "NO"
'            End If
'            UPDICU_QR(1) = ""
'            UPDICU_QR.Execute
'        End If
'    End If
    NEWFLAG = False
    'v_rep = False
End If

If MODIFYFLAG = True Then
    If txt_modcode.Text = "SP" Then
        SPUPD
    End If
'   If txt_modcode = "IH" Then  ' commented by chandu 03/01/2003
'        IHUPD
'   End If
   If txt_modcode = "FA" Then
        FAUPD
   End If
   If txt_modcode = "CG" Or txt_modcode = "CN" Then
        CGUPD
   End If
   If Trim(txt_modcode.Text) = "BA" Or Trim(txt_modcode.Text) = "MB" Or Trim(txt_modcode.Text) = "SE" Or Trim(txt_modcode.Text) = "MC" Or Trim(txt_modcode.Text) = "CM" Or Trim(txt_modcode.Text) = "MO" Or Trim(txt_modcode.Text) = "XX" Then
        MBUPD
    End If
    If Trim(txt_modcode.Text) = "ED" Then
        EDUPD
    End If
    If Trim(txt_modcode) = "2D" Or Trim(txt_modcode) = "EK" Or Trim(txt_modcode) = "PF" Or Trim(txt_modcode) = "ST" Then
        ECUPD
    End If
'    If Trim(txt_modcode.Text) = "BP" Then
'        BPASSIGN
'    End If
   MODIFYFLAG = False
End If

rdb.CommitTrans
MsgBox "Record is Successfully Saved."
btnenabl ("SAVE")
'If Trim(txt_modcode) = "2D" Or Trim(txt_modcode) = "EK" Or Trim(txt_modcode) = "PF" Or Trim(txt_modcode) = "ST" Then
'   CMB_MEMO.Enabled = True
'End If
If NoMemo = True Then
  CMB_MEMO.Enabled = False
End If
Dim STRNO As String

STRNO = req_chglen(Trim(Data1.Recordset.Fields("REQNO")))
If Trim(STRNO) = "" Then
   STRNO = Trim(cmb_reqno)
End If


'Call ASSREQNO(STRNO)
'*************ADDED BY SHWETA 19-AUG-09
'**for nc pat_catg
BMEMOEN = True 'Shraddha 01/09/09
If Trim(txt_catgory.Text) = "NC" Then
'*********updating ser_code in rrequisitn
    Dim SELDMUL_RQ As rdoQuery
    Dim SELDMUL_RS As rdoResultset
    a = Trim(STRNO)
    STRDMUL = "SELECT * FROM dISLIB.dmulproce where reqno = '" & Trim(STRNO) & "'"
    Set SELDMUL_RQ = rdb.CreateQuery("NSELDMUL_RQ", STRDMUL)
    Set SELDMUL_RS = SELDMUL_RQ.OpenResultset(1, 3)

    If SELDMUL_RS.EOF = False Then
        If SELDMUL_RS.RowCount > 1 Then
            rdb.Execute "UPDATE RISLIB.RREQUISITN SET SER_CODE ='8888' WHERE REQNNO = '" & Trim(STRNO) & "'"
        ElseIf SELDMUL_RS.RowCount = 1 Then
            rdb.Execute "UPDATE RISLIB.RREQUISITN SET SER_CODE ='" & Trim(SELDMUL_RS("PRO_CODE")) & "' WHERE REQNNO = '" & Trim(STRNO) & "'"
        End If
    End If

    If txt_status.Text = "IP" Then
        rdb.Execute "UPDATE RISLIB.RREQUISITN SET RECPTNO ='IP'  WHERE REQNNO = '" & Trim(STRNO) & "'"
    End If

    '*********p status in reqn fld
    
    Dim SELMA101_RQ As rdoQuery
    Dim SELMA101_RS As rdoResultset
    Dim STRMA101_RREQ

    STRMA101_RREQ = ""
    STRMA101_RREQ = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE = "
    STRMA101_RREQ = STRMA101_RREQ + "(SELECT ser_code  FROM RISLIB.RREQUISITN WHERE CASENO = '" & Trim(Msk_caseno.FormattedText) & "' AND REQNNO = '" & Trim(STRNO) & "') "
    STRMA101_RREQ = STRMA101_RREQ + " and NC_CH = 0"
    
    Set SELMA101_RQ = rdb.CreateQuery("SELMA101_RQ", STRMA101_RREQ)
    Set SELMA101_RS = SELMA101_RQ.OpenResultset(1, 3)
    Dim UPRRQUISTN_NC_RQ As rdoQuery
    Dim SHW

    If SELMA101_RS.EOF Then
        If txt_status.Text = "IP" Then
            BMEMOEN = True
            STRUPDT_REQSTAT = ""
            STRUPDT_REQSTAT = "UPDATE RISLIB.RREQUISITN SET REQNSTAT = NULL WHERE CASENO = ? AND REQNNO = ?"
            Set UPRRQUISTN_NC = rdb.CreateQuery("UP_REQN_STAT", STRUPDT_REQSTAT)
            UPRRQUISTN_NC(0) = Trim(Msk_caseno.FormattedText)
            UPRRQUISTN_NC(1) = Trim(STRNO)
            UPRRQUISTN_NC.Execute
        End If
    Else
        BMEMOEN = False
        STRUPDT_REQSTAT = ""
        STRUPDT_REQSTAT = "UPDATE RISLIB.RREQUISITN SET REQNSTAT = 'P' WHERE CASENO = ? AND REQNNO = ?"
        Set UPRRQUISTN_NC = rdb.CreateQuery("UP_REQN_STAT", STRUPDT_REQSTAT)
        UPRRQUISTN_NC(0) = Trim(Msk_caseno.FormattedText)
        UPRRQUISTN_NC(1) = Trim(STRNO)
        UPRRQUISTN_NC.Execute
    End If
    SELMA101_RQ.Close

End If

If SELRREQNOFLAG = False Then
    SELRREQNOQR(0) = STRNO
    Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
    SELRREQNOFLAG = True
Else
    SELRREQNOQR(0) = STRNO
    SELRREQNORS.MoreResults
    SELRREQNORS.Requery
End If
If SELRREQNORS.RowCount > 0 Then
   If Trim(SELRREQNORS("REQNSTAT")) = "P" And (Trim(txt_status.Text) = "OP") Then
      CMB_MEMO.Enabled = False
   Else
   CMB_MEMO.Enabled = BMEMOEN
   End If
End If
'************ Added by Shraddha Pandit on 22/05/2009 ***************
If Trim(txt_catgory.Text) = "NC" And Trim(txt_dept.Text) = "NZZ" And Trim(txt_modcode.Text) = "HT" Then
    'Set SEL_MEMOENABL_RS = SEL_MEMOENABL_QR.OpenResultset(1, 3)
    'cmb_dept = nzz txt_dept
    'cmb_modality = HLA txt_modcode
    HLA_MM = True
    CMB_MEMO.Enabled = True
Else
    HLA_MM = False
End If
'********************************************************************


Exit Sub
MSG:
rdb.RollbackTrans
MsgBox Err.Description
'Exit Sub
End Sub
'Sub PRO_UPREQN_STAT()
'
'
'Dim UPRRQUISTN_NC_RQ As rdoQuery
'Dim SHW
'STRUPDT_REQSTAT = ""
'STRUPDT_REQSTAT = "UPDATE RISLIB/RREQUISITN SET REQNSTAT = "
'Set UPRRQUISTN_NC = rdb.CreateQuery()
'
'End Sub

Private Sub cmd_view_Click()
btnenabl ("VIEW")
SPCLEAR
opt_patient.Value = True
'V_PROCE = False
If DATETIMEFLAG = False Then
   DATETIME
   DATETIMEFLAG = True
Else
    DATETIME
End If
txt_date.Text = DATE11
txt_time.Text = time1
lbl_reqno.Visible = True
lbl_pathno.Visible = False
cmb_reqno.Visible = True
txt_pathno.Visible = False
Msk_caseno.Enabled = True
'msk_caseno.SetFocus
viewflag = True
VIEWFLAG_BCD = True
ECMODFLAG = False
If opt_patient.Value = True Then
    Msk_caseno.Mask = "??/#####"
    'Msk_caseno.SetFocus
    
End If
If opt_staff.Value = True Then
    Msk_caseno.Mask = "######"
    Msk_caseno.SetFocus
End If
opt_patient.Value = True
End Sub

Private Sub Combo1_Change()

End Sub





Private Sub Command1_Click()
Dim VAR1
'VAR = Shell("C:\WINDOWS\Desktop\EMR 04.08.05.EXE" & Trim(Msk_caseno.FormattedText), vbNormalFocus)
VAR1 = Msk_caseno.FormattedText
VAR = Shell("C:\EMR1.EXE " & VAR1, vbNormalFocus)
AppActivate VAR

End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click() 'Print Barcode button which is invisible
If NEWFLAG_BCD = True Then
    Data1.Recordset.MoveFirst
    While Not Data1.Recordset.EOF
        'reqno = req_chglen(DAUTOREQNO)
        REQNO = req_chglen(Data1.Recordset.Fields("REQNO"))
        Barcode
        MYDELAY (1)
        'Data1.Recordset.RecordCount
        Data1.Recordset.MoveNext
    Wend
ElseIf VIEWFLAG_BCD = True Then
    REQNO = Trim(req_chglen(cmb_reqno.Text))
    Barcode
End If
End Sub
Public Sub Barcode()

'''''''''barcode BY Kanchan

Dim QTY1 As String
Dim BAR_CD As String
Dim EXP_DT As String
Dim ITEM_DESC As String
Dim TEMPDATA As String
Dim rack As String
Dim batchno1 As String
Dim itemcode1 As String
Dim I As Long, START
Dim j As Integer
Dim NOL As Integer
Dim Nos As Integer
Dim PRNTSTR As String
Dim PRNTSTR1 As String
Dim row As Integer
Dim req As String
Dim DPT As String
Dim agesex As String
Dim delay As Integer
Dim REQ111t, P_CODE111t, SCODE111t


    'If Mid(reqno, 4, 2) = "SE" Then
    If Len(Msk_caseno) = 6 Then
        f1 = FreeFile()
        Open "c:\micro.TXT" For Output As #f1
        Dim rs_emp As rdoResultset
        

        Dim DPT1 As String


        str = "SELECT distinct SAMPLECODE.SAMPLECD FROM DISLIB.SAMPLECODE,DISLIB.DMULPROCE WHERE DMULPROCE.PRO_CODE=SAMPLECODE.SERCODE AND DMULPROCE.REQNO=?"
        Set SELLABNO_QR = rdb.CreateQuery("SELLABNO", str)

        SELLABNO_QR(0) = REQNO
        SELLABNO_QR.Execute


        Set SELLABNO_RS = SELLABNO_QR.OpenResultset(1, 3)
        Dim cstaff As Integer
        cstaff = SELLABNO_RS.RowCount

                    While Not SELLABNO_RS.EOF
                        'Dim REQ111t, P_CODE111t, SCODE111t
                        REQ111t = REQNO

                        SCODE111t = SELLABNO_RS(0)
                        smplcd = SELLABNO_RS(0)

                        
                        REQ22 = Mid$(REQ111t, 4)
                        SCODE111t = REQ22 & SCODE111t
                        
                        deptdesc
                        
                        DPT1 = Trim(Mid(REQNO, 1, 3))
                        SELDEPTDES_QR(0) = DPT1
                        Set SELDEPTDES_RS = SELDEPTDES_QR.OpenResultset(1, 3)
                                                
                        DESCRIP = ""
                        DESCRIP = Mid(Trim(SELDEPTDES_RS("DEPT_NAME")), 1, 8) & "/" & colorcd
                        
                        
                        Set rs_emp = rdb.OpenResultset("select age,sex from dislib.dreqstaffd where empcode='" & Trim(Msk_caseno.FormattedText) & "'", 1, 1)
                        agesex1 = rs_emp(0) & "/" & rs_emp(1)

                        
                        PRNTSTR = "I8,1,001" & vbCrLf & "q591" & vbCrLf & "S3" & vbCrLf & "ON,D" & vbCrLf & "JF" & vbCrLf & "j189" & vbCrLf & "WN" & vbCrLf & "D10" & vbCrLf & "ZB" & vbCrLf & "Q217,44" & vbCrLf & "N" & vbCrLf
                        PRNTSTR1 = PRNTSTR & "A52,11,0,2,1,1,N,""" & Trim(txt_name) & """" & vbCrLf
                        PRNTSTR1 = PRNTSTR1 & "B92,41,0,1,3,6,79,N,""" & SCODE111t & """" & vbCrLf
                        PRNTSTR1 = PRNTSTR1 & "A194,124,0,2,1,1,N,""" & SCODE111t & """" & vbCrLf
                        PRNTSTR1 = PRNTSTR1 & "A52,152,0,3,1,1,N,""" & agesex1 & """" & vbCrLf
                        PRNTSTR1 = PRNTSTR1 & "A150,150,0,4,1,1,N,""" & Trim(Msk_caseno) & """" & vbCrLf
                        PRNTSTR1 = PRNTSTR1 & "A340,149,0,2,1,1,N,""" & DESCRIP & """" & vbCrLf
                        
                        PRNTSTR1 = PRNTSTR1 & "P1" & vbCrLf
                        PRNTSTR1 = PRNTSTR1 & "N" & vbCrLf
                        Print #f1, PRNTSTR1




                        SELLABNO_RS.MoveNext
                    Wend
        Close #f1
        Shell "c:\DISREPORTSEC\DATA1\micro.bat", 0
        'SELPROCODE_QR.Close
        SELLABNO_QR.Close
        'Exit Sub
    Else
        'Kill "C:\MICRO.TXT"
        f1 = FreeFile()
        Open "c:\micro.TXT" For Output As #f1
        Dim rs_ta101 As rdoResultset



        str = "SELECT distinct SAMPLECODE.SAMPLECD FROM DISLIB.SAMPLECODE,DISLIB.DMULPROCE WHERE DMULPROCE.PRO_CODE=SAMPLECODE.SERCODE AND DMULPROCE.REQNO=?"
        Set SELLABNO_QR = rdb.CreateQuery("SELLABNO", str)

        SELLABNO_QR(0) = REQNO
        SELLABNO_QR.Execute
        
        Set SELLABNO_RS = SELLABNO_QR.OpenResultset(1, 3)
        Dim c As Integer
        c = SELLABNO_RS.RowCount

                    While Not SELLABNO_RS.EOF
                        'Dim REQ111t, P_CODE111t, SCODE111t
                        REQ111t = REQNO

                        SCODE111t = SELLABNO_RS(0)
                        smplcd = SELLABNO_RS(0)
                        
                        REQ22 = Mid$(REQ111t, 4)
                        SCODE111t = REQ22 & SCODE111t

                                    
                        deptdesc
                        DPT = Trim(Mid(REQNO, 1, 3))
                        
                        SELDEPTDES_QR(0) = DPT
                        Set SELDEPTDES_RS = SELDEPTDES_QR.OpenResultset(1, 3)
                        
                        DESCRIP = ""
                        DESCRIP = Mid(Trim(SELDEPTDES_RS("DEPT_NAME")), 1, 8) & "/" & colorcd
                        
                        Set rs_ta101 = rdb.OpenResultset("select birth_date,sex from patoolslib.ta101 where case_no='" & Trim(Msk_caseno.FormattedText) & "'", 1, 1)

                        Dim n As Integer
                        n = DateDiff("m", rs_ta101("BIRTH_DATE"), CStr(DATE11)) / 12
                        agesex = n & "/" & rs_ta101(1)
                                                                      
                        
                        PRNTSTR = "I8,1,001" & vbCrLf & "q591" & vbCrLf & "S3" & vbCrLf & "ON,D" & vbCrLf & "JF" & vbCrLf & "j189" & vbCrLf & "WN" & vbCrLf & "D10" & vbCrLf & "ZB" & vbCrLf & "Q217,44" & vbCrLf & "N" & vbCrLf
                        
                        PRNTSTR1 = PRNTSTR & "A52,11,0,2,1,1,N,""" & Trim(txt_name) & """" & vbCrLf
                        PRNTSTR1 = PRNTSTR1 & "B92,41,0,1,3,6,79,N,""" & SCODE111t & """" & vbCrLf
                        PRNTSTR1 = PRNTSTR1 & "A194,124,0,2,1,1,N,""" & SCODE111t & """" & vbCrLf
                        PRNTSTR1 = PRNTSTR1 & "A52,152,0,3,1,1,N,""" & agesex & """" & vbCrLf
                        PRNTSTR1 = PRNTSTR1 & "A150,150,0,4,1,1,N,""" & Trim(Msk_caseno) & """" & vbCrLf
                        PRNTSTR1 = PRNTSTR1 & "A340,149,0,2,1,1,N,""" & DESCRIP & """" & vbCrLf
                        PRNTSTR1 = PRNTSTR1 & "P1" & vbCrLf
                        PRNTSTR1 = PRNTSTR1 & "N" & vbCrLf
                        Print #f1, PRNTSTR1
                        
                        SELLABNO_RS.MoveNext
                    
                    Wend
        Close #f1
        Shell "c:\DISREPORTSEC\DATA1\micro.bat", 0
        'SELPROCODE_QR.Close
        SELLABNO_QR.Close
        'Exit Sub
    End If


  '  End If



End Sub
Public Sub deptdesc()

'''' Done by kanchan to select color for printing on the barcode label

colorcd = ""
    If smplcd = "A" Or smplcd = "B" Then
        colorcd = "RED"
    ElseIf smplcd = "R" Or smplcd = "S" Or smplcd = "P" Then
        colorcd = "GRAY"
    ElseIf smplcd = "U" Then
        colorcd = "BLUE"
    ElseIf smplcd = "C" Then
        colorcd = "PURPLE"
    ElseIf smplcd = "D" Then
        colorcd = "BLACK"
    Else
        colorcd = ""
    End If
End Sub
Private Sub Form_Load()
'On Error GoTo MSG:
BMEMOEN = True
'added by milind on 17/04/06 to get a link from ot module
    
    VAR = COMMAND()
    If Len(VAR) > 0 Then
        UID = "A" + "DISCLK"
        If Mid(Trim(UID), 2) = "DISCLK" Then
            Main
        End If
    End If
    
'upto here
    
'************** Added by Shraddha Pandit on 29/04/2009 *****************

Dim sample_QR As rdoQuery
Dim sample_RS As rdoResultset
Dim sql As String

sql = "SELECT * FROM DISLIB.microsampl order by SAMPLENAME"
Set sample_QR = rdb.CreateQuery("", sql)
Set sample_RS = sample_QR.OpenResultset(1, 3)

If Not sample_RS.EOF Then
      
    While Not sample_RS.EOF
        
        cmb_specimen.AddItem sample_RS("SAMPLENAME")
        sample_RS.MoveNext
        
    Wend
        
End If
sample_RS.Close

'********
HLA_MM = False
'***********************************************************
    
    
NEWFLAG = False
AMFLAG = False
viewflag = False
MODIFYFLAG = False
IPOPMEMO = False
'INNC = False
For I = 1 To 11
CMB_DURILLMONTH.AddItem I
Next I
CMB_TEXT.clear
cmb_frozensec.AddItem "YES"
cmb_frozensec.AddItem "NO"
' FOR RTIMELOG
str = "INSERT INTO RISLIB.RTIMELOG(REQNNO,CASENO,REQNDATE,REQNTIME) VALUES(?,?,?,?)"
Set INSTIME_QR = rdb.CreateQuery("", str)
'********** for Selecting Deprtment ***********
    str = "SELECT * FROM DISLIB.DISLNKUSR WHERE USR_ID = ? "
    Set sel_DptCd_qr = rdb.CreateQuery("sel_DptCd_qr", str)
    sel_DptCd_qr(0) = UID
    Set sel_DptCd_rs = sel_DptCd_qr.OpenResultset(1, 3)
    Dept_cd1 = ""
    If sel_DptCd_rs.RowCount > 0 Then
      Dept_cd1 = Trim(sel_DptCd_rs("DEPT_CODE"))
      enter_by1 = Trim(sel_DptCd_rs("USERNAME"))
    End If
    sel_DptCd_qr.Close
    'MsgBox Dept_cd1
    

'*******************for selectin provisional diagnosis******************
str = "select prvdiagdes from patoolslib.ta1012 where case_no=?"
Set selprvdgs_qr = rdb.CreateQuery("", str)
'*****

If pathologyFLAG = False Then
    
    DATETIME
'    DATE11 = Format(DATE11, "DD/MM/YYYY")
    txt_date.Text = DATE11
    txt_time.Text = time1
    
    'molecular pathology
    str = "SELECT * FROM DISLIB.DREQUISITN WHERE CASENO=? ORDER BY REQNO "
    Set SELDREQMPQR = rdb.CreateQuery("SELDREQQRmp", str)
   
    str = "INSERT INTO DISLIB.DMOLEPROC(REQNO,PRO_CODE,REQNDT,CASENO) VALUES(?,?,?,?)"
    Set INSDMOLPROC_QR = rdb.CreateQuery("INSDMOLEROCE_QRMP", str)
    
    str = "INSERT INTO DISLIB.DMOLEREQ(REQNO,CASENO,SITE,LATERALITY,FRESH,ARCHIVE) VALUES(?,?,?,?,?,?)"
    Set INSDMOLREQ_QR = rdb.CreateQuery("INSDMOLREQ_QRMP", str)
    
    '*** inserting into TBPRODUCT of Penpol
        str = "Insert into dislib.TBProduct(FCaseNo,FREQNNO,FCOMPONENT,FQTY,FREMARKS,FSPLREQ,FPHLEB,FREPLCDNO,FISSUED,FFLAG) values(?,?,?,?,?,?,?,?,?,?)"
        Set Ins_TBProduct_qr = rdb.CreateQuery("Ins_TBProduct_qr", str)
    '*** inserting into TBPRODUCT of Penpol
    
    ''***********BP CROSSMATCH***************
    str = "UPDATE RISLIB.RREQUISITN SET REQNSTAT=NULL WHERE REQNNO=?"
    Set UPDREQNSTAT_QR = rdb.CreateQuery("UPDREQNSTAT_STR", str)
    
    str = "UPDATE DISLIB.DMULPROCE SET COLLECTION=? WHERE  REQNO=? AND PRO_CODE='H002'"
    Set UPDDMULPROCE_QR = rdb.CreateQuery("UPDDMULPROCE_STR", str)
    
    str = "SELECT * FROM DISLIB.PRFREECODE WHERE CODE=?"
    Set PRFREE_QR = rdb.CreateQuery("PRFREECODE", str)
    PRFREE_FLG = False
    
    ''***********BP CROSSMATCH***************
    
    '************FOR BP(RDP,SDP,FFP,FVIIID,CRYO,ALB) UPDATE MEMO NO AS NOCHARGE **************
      
      str = "UPDATE RISLIB.RREQUISITN SET MEMONO=?,MEMODT=? WHERE REQNNO=? AND CASENO=?"
      Set UPDMEMORREQ_QR = rdb.CreateQuery("UPDMEMORREQ_QR", str)
      
    '************FOR BP(RDP,SDP,FFP,FVIIID,CRYO,ALB) UPDATE MEMO NO AS NOCHARGE **************
    '' FOR CYTOGENITICS QUERIES 13-04-2005
    str = "INSERT INTO DISLIB.CYTOGCLDTA (REQNO,CASEN0,AT_DIAG,PR_TREAT,PS_INDU,RELAPSE,REMISSION,FOLLUP,POSTBMT,B_SEX,SP_TYPE,CGDIAG1,CGDIAG2,FAX,PHONE,REMARK1,REMARK2)" _
         & "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
    Set INSCYTOGCLDTACY_QR = rdb.CreateQuery("INSERTCYTGCL_QR1", str)
    
    str = "UPDATE DISLIB.CYTOGCLDTA CASEN0=?,AT_DIAG=?,PR_TREAT=?,PS_INDU=?,RELAPSE=?,REMISSIO=?,FOLLUP=?,POSTBMT=?,B_SEX=?,SP_TYPE=? WHERE REQNO=?"
    Set UPDCYTOGCLDTACY_QR = rdb.CreateQuery("UPDERTCYTGCL_QR1", str)

    str = "SELECT * FROM DISLIB.CYTOGCLDTA WHERE REQNO=?"
    Set SELCYTOGCLDTACY_QR = rdb.CreateQuery("SELRREQUISITN_QRcy", str)
    
    str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE=? AND SER_CODE<>'F106' ORDER BY SER_DET"
    Set SELMA101CY_QR = rdb.CreateQuery("SELMA101CY_QRrr11cyc11c1", str)
       
    str = "SELECT * FROM RISLIB.RINVESTIGN where modalitycd=?"
    Set SELRINVESTCY_QR = rdb.CreateQuery("slerinevetcy11c1c1", str)
    
    str = "select b.modlitydes from rislib.rrequisitn a,rislib.rmodality b  where a.modalitycd=b.modalitycd and a.reqnno=?"
    Set selreqmoda_qr = rdb.CreateQuery("selmodareqc11c", str)
    
    str = "Select reqnno from rislib.rrequisitn where (dept_code='UZZ' or dept_code='FZZ' or dept_code='GZZ' or dept_code='VZZ') and caseno=?"
    Set selreqnno_qr = rdb.CreateQuery("select_qrc11c", str)
  
    selreqnno_qr(0) = "TT/00500"
    Set selreqnno_rs = selreqnno_qr.OpenResultset(2, 3)
    
    
   '' FOR CYTOGENITICS QUERIES 13-04-2005
    
    ''**********HAEMATOLOGY PREPATHNO************
    str = "SELECT REQNNO,REQNDT,REQNSTAT,CASENO,PREPORTDT,PATHNO FROM DISLIB.DSPLF WHERE CASENO=? AND (REQNSTAT='T' OR REQNSTAT='F') AND (REQNNO LIKE 'FZZCG%' OR REQNNO LIKE 'FZZSP%' OR REQNNO LIKE 'FZZCN%' OR  REQNNO LIKE 'FZZFA%' OR REQNNO LIKE 'UZZ%') ORDER BY REQNNO"
    Set PREPATH_QR = rdb.CreateQuery("PREPATH_STR", str)
    
    str = "INSERT INTO DISLIB.DHEMCYTO1 (REQNO,MPO,SBB,PAS,LAP,NSE,TRAP,cbc,hem,cg,cyto,smar,fld1,fld2,fld3,fld4) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
    Set INSDHEMCYTO_QR = rdb.CreateQuery("INSDHEMCYTO1_STR", str)
    
    str = "SELECT REQNO,MPO,SBB,PAS,LAP,NSE,TRAP,cbc,hem,cg,cyto,smar,fld1,fld2,fld3,fld4 FROM    DISLIB.DHEMCYTO1  WHERE REQNO=?"
    Set SELDHEMCYTO_QR1 = rdb.CreateQuery("SELCYTO11x_STR", str)
    
    
    ''**********HAEMATOLOGY PREPATHNO************

    
    
    ''************ICU & RECOVERY WARD**************
    
    str = "UPDATE DISLIB.DREQUISITN SET ICDS2=? WHERE REQNO=?"
    Set UPDICU_QR = rdb.CreateQuery("UPDICU_STR", str)
    
    str = "SELECT ICDS2 FROM DISLIB.DREQUISITN WHERE REQNO=?"
    Set SELICU_QR = rdb.CreateQuery("SELICU_STR", str)
    
    str = "SELECT * FROM DISLIB.DISICU WHERE IPADD=?"
    Set SELDISICU_QR = rdb.CreateQuery("SELDISICU_STR", str)
    
    '''************ICU*************
    
    ''**********HL*****************
    str = "SELECT * FROM DISLIB.DBMACIND WHERE REQNO=?"
    Set SELDBMACIND_QR = rdb.CreateQuery("SELDBMACIND_STR", str)
    
    ''**********HL*****************
    
'************ Shraddha Pandit 21/05/2009 ***************
str = "SELECT * FROM PATOOLSLIB.MA101 WHERE NC_CH>0 and SER_CODE like 'H%' "
Set SEL_MEMOENABL_QR = rdb.CreateQuery("", str)
'*******************************************************
    
    
    ''************FOR PREPATHNO & REQNO COMBOBOX REPORT ***********
    
    'STR = "SELECT * FROM RISLIB.RREQUISITN LEFT OUTER JOIN DISLIB.DREQUISITN ON RISLIB.RREQUISITN.REQNNO=DISLIB.DREQUISITN.REQNO WHERE RISLIB.RREQUISITN.CASENO=? AND (REQNSTAT='T' OR REQNSTAT='F') AND (REQNNO LIKE 'FZZCG%' OR REQNNO LIKE 'FZZSP%' OR REQNNO LIKE 'FZZCN%' OR  REQNNO LIKE 'FZZFA%' OR REQNNO LIKE 'LZZCR%' OR REQNNO LIKE 'LZZCT%' OR REQNNO LIKE 'LZZMM%' OR REQNNO LIKE 'LZZMR%' OR REQNNO LIKE 'LZZUS%')" 'ORDER BY RREQUISITN.REQNDT DESC"
    str = "SELECT REQNNO,REQNDT,REQNSTAT,CASENO,PREPORTDT,PATHNO FROM DISLIB.DSPLF WHERE CASENO=? AND (REQNSTAT='T' OR REQNSTAT='F') AND (REQNNO LIKE 'FZZCG%' OR REQNNO LIKE 'FZZSP%' OR REQNNO LIKE 'FZZCN%' OR  REQNNO LIKE 'FZZFA%' OR REQNNO LIKE 'LZZCR%' OR REQNNO LIKE 'LZZCT%' OR REQNNO LIKE 'LZZMM%' OR REQNNO LIKE 'LZZMR%' OR REQNNO LIKE 'LZZUS%') ORDER BY PREPORTDT"
    Set seldrequisitn_qr10 = rdb.CreateQuery("seldrequisitn11_1qr110", str)
         
    str = "SELECT REQNNO,REQNDT,REQNSTAT,CASENO,PREPORTDT,PATHNO FROM DISLIB.DSPLF WHERE CASENO=? AND REQNSTAT='F' AND (REQNNO LIKE 'FZZCG%' OR REQNNO LIKE 'FZZSP%' OR REQNNO LIKE 'FZZCN%' OR  REQNNO LIKE 'FZZFA%' OR REQNNO LIKE 'LZZCR%' OR REQNNO LIKE 'LZZCT%' OR REQNNO LIKE 'LZZMM%' OR REQNNO LIKE 'LZZMR%' OR REQNNO LIKE 'LZZUS%')"
    Set seldrequisitn_qr11 = rdb.CreateQuery("seldrequisitn11_1qrREP", str)
     
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
    Set SELRREQNOQR = rdb.CreateQuery("SELRREQNOQR1SPREQ", str)
     
    str = "SELECT  * FROM DISLIB.DADDENDUM WHERE REQNO=?"
    Set SELDADDENDUM_QR = rdb.CreateQuery("selDADDENDUM_QR11SPREQ", str)
       
    str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO =? "
    Set SELPATHDREQUISITN_QR = rdb.CreateQuery("SELPATHDREQUISITN1SPREQ_QR1", str)
    
    str = "SELECT REQNO,FSPREPORT,IHREPORT,IMPRESSION,GROSSDESC,GROSSEDBY,SPREPORT,REPORTIND,FNAC FROM DISLIB.DFREPORT WHERE REQNO=?"
    Set SELDFREPORT_QR = rdb.CreateQuery("SELDFREPORT_1QR1sprep", str)
    
    
    ''************FOR PREPATHNO & REQNO COMBOBOX REPORT ***********
    ' THIS QUERY CHANGED FOR TESTING PRINTING
    'STR = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=? AND PAT_CATG NOT IN('NC','SF')AND PATNSTAT<>'ST'"
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=? "
    Set SELRISSTAT_QR = rdb.CreateQuery("SELRISSTAT_STR", str)
    
    '''check for zero amount
    str = "select * from patoolslib.ma101 where ser_code=?"
    Set SELMA101hl_QR = rdb.CreateQuery("SELMA101hl_str", str)
    
   
    
    '''update requisition table if charge is zero
    
    str = "update rislib.rrequisitn set reqnstat='P' where reqnno=?"
    Set updreq_qr = rdb.CreateQuery("updreq_str", str) 'PAM
    
    
    'STR = "SELECT DISTINCT(*) FROM DISLIB.DTMTAB WHERE CASENO=?"
    'Set SELDTMTAB_QR = rdb.CreateQuery("SELDTMTAB_STR1", STR)
    
    ''INSERT INTO DTMNO TABLE
    str = "INSERT INTO DISLIB.DTMNO(CASENO,DTMNO,FLD1) VALUES(?,?,?)"
    Set INSDTMNO_QR = rdb.CreateQuery("INSDTMNO_STR", str)
    
    
    str = "select * from PATOOLSLIB.ta101 where case_no=?"
    Set SELTA101_QR = rdb.CreateQuery("1selta101_qr1", str)
    ' lzz rEMOVED
    'STR = "SELECT * FROM PATOOLSLIB.MA106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ' OR DEPT_CODE='CAZ' OR DEPT_CODE='CBZ' OR DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='IZZ' OR DEPT_CODE='VZZ') "
    str = "SELECT * FROM PATOOLSLIB.MA106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ' OR DEPT_CODE='CAZ' OR DEPT_CODE='CBZ' OR DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='VZZ') "
    Set SELMA106_QR = rdb.CreateQuery("", str)
    
    str = "SELECT * FROM PATOOLSLIB.MA10121"
    Set SELMA10121_QR = rdb.CreateQuery("SELMA101211_QR11", str)

    str = "select * from PATOOLSLIB.ma104 where case_no=?"
    Set selma104_qr = rdb.CreateQuery("selma1041_qr1", str)

    str = "select * from PATOOLSLIB.ta1011 where case_no=?"
    Set SELTA1011_QR = rdb.CreateQuery("selta1011_qr1", str)

    str = "SELECT * FROM RISLIB.CLINHIST WHERE CASENO=? ORDER BY CLINHISTDT,CLINHISTTM "
    Set SELCLINHIST_QR = rdb.CreateQuery("SELCLINHIST_QR1", str)

    str = "select * from rislib.rmodality where dept_code=?"
    Set SELRMODALITY_QR = rdb.CreateQuery("selrmodality_qr11", str)
    
    str = "SELECT * FROM RISLIB.RINVESTIGN WHERE MODALITYCD=?"
    Set SELRINVEST_QR = rdb.CreateQuery("SELCLINHIST_QR11", str)
    
    ''ADD 30/05/03
    
    str = "SELECT * FROM RISLIB.RINVESTIGN WHERE MODALITYCD=? AND SER_CODE<>'F106'"
    Set SELRINVEST1_QR = rdb.CreateQuery("SELCLINHIST_QR111", str)
    
    str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE=? AND SER_CODE<>'F106' ORDER BY SER_DET"
    Set SELMA101_QR1 = rdb.CreateQuery("SELMA101_QRrr11", str)
    
    ''------------
    
    ''ADD ON 02/05/2003 BY MONO
    
    str = "SELECT * FROM RISLIB.RINVESTIGN WHERE MODALITYCD=? AND SER_CODE<>'F406'"
    Set SELRINVEST11_QR = rdb.CreateQuery("SELCLINHIST_QR1111", str)
    
    str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE=? AND SER_CODE<>'F406'"
    Set SELMA101_QR11 = rdb.CreateQuery("SELMA101_QRrr111", str)
    
    ''************
    str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE=?"
    Set SELMA101_QR = rdb.CreateQuery("SELMA101_QRrr1", str)
    
    str = "SELECT * FROM PATOOLSLIB.MA301 ORDER BY SITEDESC "
    Set SELMA301_QR = rdb.CreateQuery("SELMA301_QR1", str)
    
    str = "SELECT * FROM PATOOLSLIB.MA10122"
    Set SELMA10122_QR = rdb.CreateQuery("SELMA10122_QR1", str)
    
    str = "SELECT CESSYN FROM EMPLIB.ELEMPINFO WHERE EMPCODE=?"
    Set SELELEMP_QR = rdb.CreateQuery("SELELEMP_QR", str)
    
    str = "SELECT * FROM EMPLIB.EMPBASINFO WHERE EMPCODE=?"
    Set SELEMPBASINFO_QR = rdb.CreateQuery("SELEMPBASINFO_QR1111", str)
    
    str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=?"
    Set SELSTAFFINFO_QR = rdb.CreateQuery("SELSTAFFINFO_QR1111", str)
    SELSTAFFINFOFLAG = False
    
    '***** AUTOGENERATION OF THE REQUISITION NO.
    str = "SELECT * FROM DISLIB.DAUTOREQNO WHERE DEPT_CODE=? AND MODALITYCD=? AND REQNYEAR=?"
    Set SELDAUTOREQNO_QR = rdb.CreateQuery("SELDAUTOREQNO_QR1", str)
    
    str = "INSERT INTO DISLIB.DAUTOREQNO (DEPT_CODE,MODALITYCD,REQNYEAR,REQNSRNO) VALUES(?,?,?,?)"
    Set INSDAUTOREQNO_QR = rdb.CreateQuery("INSDAUTOREQNO_QR1", str)
    
    str = "UPDATE DISLIB.DAUTOREQNO SET REQNSRNO=? WHERE DEPT_CODE=? AND MODALITYCD=? AND REQNYEAR=?"
    Set UPDDAUTOREQNO_QR = rdb.CreateQuery("UPDDAUTOREQNO_QR1", str)
    ' ***** AUTOGENERATION OF THE REQUISTION NO. QURIES
    
    ' ***** AUTOGENERATION OF THE PATHNO.
    str = "SELECT * FROM DISLIB.DAUTOPATH WHERE DEPT_CODE=? AND MODALITYCD=? AND PATHYEAR1=?"
    Set SELDAUTOPATH_QR = rdb.CreateQuery("SELDAUTOPATH_QR1", str)
    
    str = "INSERT INTO DISLIB.DAUTOPATH (DEPT_CODE,MODALITYCD,PATHYEAR,PATHYEAR1,PATHNO) VALUES (?,?,?,?,?)"
    Set INSDAUTOPATH_QR = rdb.CreateQuery("INSDAUTOPATH_QR1", str)
    
    str = "UPDATE DISLIB.DAUTOPATH SET PATHNO=? WHERE DEPT_CODE=? AND MODALITYCD=? AND PATHYEAR1=?"
    Set UPDDAUTOPATH_QR = rdb.CreateQuery("UPDDAUTOPATH_QR1", str)
        
    ' ***** AUTOGENERATION OF THE PATH NO.
   str = "INSERT INTO RISLIB.RREQUISITN (REQNNO, REQNDT,REQNTM,CASENO,DEPT_CODE,MODALITYCD,PATNSTAT,PAT_CATG,SER_CODE,REFBYDOC,PREVEXAM,RECPTNO,RECPTDT,REQNSTAT,RECORDNO,APNTISSDT) VALUES " _
        & "(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
   Set INSRREQUISITN_QR = rdb.CreateQuery("INSRREQUISITN_QR1", str)

     
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? ORDER BY MODALITYCD,REQNNO DESC"
    Set SELRREQUISITN_QR = rdb.CreateQuery("SELRREQUISITN_QR11111", str)
    
    str = "INSERT INTO DISLIB.DREQUISITN(REQNO,PATHNO,CASENO,SITE,LATERALITS,DURILLNESS,METASTASIS,LATERALITM,OUTSSPECIM,OUTBLOCK,OUTSLIDES,FROZENSEC,FROZENDATE,IMMHISCHE,REFDOC,REFTODOC,OPPROCEDUR,WARD_NAME,BEDNO,PREPATHNO,NMETERIALS,labno)VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
    Set INSDREQUISITN_QR = rdb.CreateQuery("INSDREQUISITN_QR1", str)
    
    str = "UPDATE PATOOLSLIB.TA1012 SET PRVDIAGDES=? WHERE CASE_NO=?"
    Set UPDTA1012_QR = rdb.CreateQuery("UPDTA1012_QR1", str)
    
    str = "INSERT INTO RISLIB.CLINHIST(CASENO,CLINHISTDT,CLINHISTTM,CLINHIST) VALUES (?,?,?,?)"
    Set insclinhist_qr = rdb.CreateQuery("insclinhist_qr1", str)
    
    str = "UPDATE RISLIB.CLINHIST SET CLINHIST=? WHERE CASENO=?"
    Set updclinhist_qr = rdb.CreateQuery("updclinhist_qr1", str)
    
    str = "SELECT * FROM PATOOLSLIB.TA1012 WHERE CASE_NO=?"
    Set selta1012_qr = rdb.CreateQuery("SELTA1012_QR11", str)
    
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=? "
    Set SELRREQUISITN_QR2 = rdb.CreateQuery("SELRREQUISITN_QR21", str)
    
    str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO=?"
    Set seldrequisitn_qr = rdb.CreateQuery("SELDREQUISITN_QR1111", str)
    
    'STR = "SELECT * FROM DISLIB.DSPLF WHERE CASENO=? AND (REQNSTAT='F' or REQNSTAT='T') AND (REQNNO LIKE 'FZZCG%' OR REQNNO LIKE 'FZZCN%' OR  REQNNO LIKE 'FZZFA%' ) ORDER BY PREPORTDT DESC"
    str = "SELECT * FROM DISLIB.DREQUISITN WHERE CASENO=? ORDER BY REQNO"
    Set SELDREQQR = rdb.CreateQuery("SELDREQQR1", str)
    
    str = "SELECT SURGPROCNO FROM OTLIB.OTTSRGRECN WHERE CASENO=? ORDER BY SCHNO DESC"
    Set SELSCOPYFNDQR = rdb.CreateQuery("SELSCOPYFNDQR", str)
    
    str = "SELECT * FROM RISLIB.RMODALITY WHERE MODALITYCD=?"
    Set SELRMODALITY_QR2 = rdb.CreateQuery("SELRMODALTY_QR3", str)
    
    'STR = "SELECT * FROM PATOOLSLIB.MA106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ' OR DEPT_CODE='CAZ' OR DEPT_CODE='CBZ' OR DEPT_CODE='CCC' OR DEPT_CODE='NZZ' OR DEPT_CODE='LZZ')"
    str = "SELECT * FROM PATOOLSLIB.MA106 WHERE DEPT_CODE=? "
    Set SELMA106_QR2 = rdb.CreateQuery("SELMA106_QR3", str)
    
    str = "INSERT INTO DISLIB.DREQUISITN(REQNO,PATHNO,CASENO,OUTBLOCK,OUTSLIDES,IMMHISCHE,REFDOC,REFTODOC,WARD_NAME,BEDNO,DURILLNESS,SPIHREQNO,NMETERIALS) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)"
    Set INSIHREQ_QR = rdb.CreateQuery("INSIHREQ_QR1", str)
    
    str = "SELECT * FROM PATOOLSLIB.TA113 WHERE CASE_NO=?"
    Set SELTA113_QR = rdb.CreateQuery("SELTA113_QR1", str)
    
    str = "SELECT * FROM PATOOLSLIB.TA114 WHERE CASE_NO=?"
    Set SELTA114_QR = rdb.CreateQuery("SELTA114_QR1", str)
    
    str = "SELECT * FROM DISLIB.DPATHYEAR WHERE AYEAR=?"
    Set DPATHYEAR_QR = rdb.CreateQuery("DPATHYEAR_QR1", str)
    
    str = "INSERT INTO DISLIB.DMULPROCE(REQNO,PRO_CODE,REQNDT,CASENO) VALUES(?,?,?,?)"
    Set INSDMULPROCE_QR = rdb.CreateQuery("INSDMULPROCE_QR1", str)
    
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE DEPT_CODE=? AND modalitycd=? and caseno=? ORDER BY REQNDT DESC"
    Set SELRREQ_QR = rdb.CreateQuery("selrreq_qr1", str)
    
    
    str = "SELECT * FROM PATOOLSLIB.MA301 WHERE SITECD=?"
    Set SELMA301QR = rdb.CreateQuery("SELMA301QR1", str)
    
   'STR = "SELECT * FROM EMPLIB.ELEMPINFO WHERE (EMPCODE LIKE '1%' OR EMPCODE LIKE '3%') AND DEPTCODE='FZZ' AND CESSYN='4'"
   str = "SELECT DISTINCT UD_NAME FROM PATOOLSLIB.MA1031"
   Set SELMA1031_QR = rdb.CreateQuery("SELMA1031_QR1", str)
   'FOLLOWING LINE COMMENTED AND ADDED BY SHWETA 20-07-2012
   'str = "SELECT * FROM EMPLIB.EMPDPDLIST WHERE EMPCODE=?"
   str = "SELECT * FROM EMPLIB.EMPDPDLIST WHERE EMPCODE=? AND OPTCHSSYN ='Y'"
   Set SELEMPDPDLIST_QR = rdb.CreateQuery("SELEMPDPDLIST_QR1", str)
    
  str = "INSERT INTO DISLIB.DREQSTAFFD(REQNO,EMPCODE,PNAME,AGE,SEX,RELATION) VALUES (?,?,?,?,?,?)"
  Set INSDREQSTAFFD_QR = rdb.CreateQuery("INSDREQSTAFFD_QR1", str)
    
  str = "SELECT * FROM EMPLIB.EMPDPDLIST WHERE EMPCODE=? AND DPDREL=?"
  Set SELEMPDPDLIST1_QR = rdb.CreateQuery("SELEMPDPDLIST1_QR1", str)
  
  str = "INSERT INTO DISLIB.DIHREPORT(reqno,testorder) values (?,?)"
  Set INSDIHREPORT_QR = rdb.CreateQuery("insdihreport1_qr1", str)
    
 str = "SELECT * FROM DISLIB.DIHREPORT WHERE REQNO=?"
 Set seldihreport_qr = rdb.CreateQuery("SELDIHREPORT1_QR", str)
 
 str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE LIKE 'O%'"
 Set selMA101QR = rdb.CreateQuery("SELMA101QR1", str)
 
 str = "INSERT INTO PATOOLSLIB.TA1012(CASE_NO,PRVDIAGDES) VALUES (?,?)"
 Set insta1012_qr = rdb.CreateQuery("selta1012_1qr1", str)
 
 str = "select * from patoolslib.ma301 where sitedesc = ?"
 Set sitedupqr = rdb.CreateQuery("site100", str)
 
 str = "select * from patoolslib.ma301 where sitecd like '5%' order by sitecd desc"
 Set sitedupqr1 = rdb.CreateQuery("site1", str)
 
 str = "insert into patoolslib.ma301(sitecd,sitedesc) values(?,?)"
 Set siteinsertqr = rdb.CreateQuery("siteInsertdup1", str)
 
 
 str = "UPDATE DISLIB.DREQUISITN SET REFDOC=?,REFTODOC=?,DURILLNESS=?,SITE=?,LATERALITs=?,OPPROCEDUR=?,NMETERIALS=?,METASTASIS=?,LATERALITM=?,OUTSLIDES=?,OUTSSPECIM=?,OUTBLOCK=?,FROZENSEC=?,FROZENDATE=?,LABNO=? WHERE REQNO=?"
 Set upddrequisitn_qr = rdb.CreateQuery("UPDDREQUISITN1_QR1", str)
    
 str = "UPDATE DISLIB.DREQUISITN SET OUTSLIDES=?,OUTBLOCK=?,NMETERIALS=? WHERE REQNO=?"
 Set UPDDREQIHQR = rdb.CreateQuery("UPDREQIH1QR", str)
 
 str = "UPDATE RISLIB.RREQUISITN SET REFBYDOC=? WHERE REQNNO=?"
 Set updrrequisitn_qr = rdb.CreateQuery("updrrequisi1tn_qr1", str)
 
 str = "SELECT * FROM DISLIB.DSFTDCY WHERE REQNO=?"
 Set seldsftdcyqr = rdb.CreateQuery("seldsftdcyqrCN", str)
 
 str = "INSERT INTO DISLIB.DSFTDCY(REQNO,SPFIND,TRDET,OUTSPEC) VALUES(?,?,?,?)"
 Set insdsftdcyqr = rdb.CreateQuery("insdsftdcyqrCY", str)
 
 str = "UPDATE DISLIB.DSFTDCY SET SPFIND=?,TRDET=?,OUTSPEC=? WHERE REQNO=?"
 Set upddsftdcyqr = rdb.CreateQuery("upddsftdcyqrCN", str)
 
 str = "SELECT * FROM DISLIB.DISSEC WHERE ADDENDUMBY='YES' AND MODALITY='SP'"
 Set SELDENTERQR = rdb.CreateQuery("SELDENTERQRFN", str)
 
 str = "UPDATE RISLIB.RREQUISITN set RECPTNO=?,RECPTDT=?,REQNSTAT=?,RECORDNO=? WHERE REQNNO=?"
 Set UPDRREQQR = rdb.CreateQuery("UPDRREQQRSP", str)
 
 str = "INSERT INTO DISLIB.DMSPONANTI(REQNO,SPECIMEN,ONGOINGANT) VALUES(?,?,?)"
 Set INSSPONANTIBQR = rdb.CreateQuery("INSSPONANTIBQRMB", str)
 
 str = "UPDATE DISLIB.DMSPONANTI SET SPECIMEN=?,ONGOINGANT=? WHERE REQNO=?"
 Set UPDSPONANTIBQR = rdb.CreateQuery("UPDSPONANTIBQRMB", str)
 
 str = "SELECT * FrOM DISLIB.DMSPONANTI WHERE REQNO=?"
 Set SELSPONANTIBQR = rdb.CreateQuery("SELSPONANTIBQRMB1", str)
 
 str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=?"
 Set SELDMULPROCEQR = rdb.CreateQuery("SELDMULPROCEQRMB", str)
 
 str = "SELECT * FROM PATOOLSLIB.TA101 WHERE CASE_NO =?  AND  (DAYS(CURRENT DATE) - DAYS(REG_DATE)) <= 180"
 Set PRE_QUERY = rdb.CreateQuery("PRE_QUERYSCA", str)
 
 str = "INSERT INTO DISLIB.INDICATION(REQNO,IND1,IND2,IND3,IND4,IND5,IND6,IND7,IND8,IND9,TP1,TP2,TP3,DRUG,DOSE,NOOFCYCLES,BSA,OTHERS) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
 'STR = "INSERT INTO DISLIB.INDICATION(REQNO,IND1,IND2,IND3,IND4,IND5,IND6,IND7,IND8,IND9) VALUES (?,?,?,?,?,?,?,?,?,?)"
 Set INSINDICATIONQR = rdb.CreateQuery("INSINDICATIONQR1", str)
 
 str = "SELECT * FROM DISLIB.INDICATION WHERE REQNO=?"
 Set SELINDICATIONQR = rdb.CreateQuery("SELINDICATIONQR1", str)
 
 str = "UPDATE DISLIB.INDICATION SET IND1=?,IND2=?,IND3=?,IND4=?,IND5=?,IND6=?,IND7=?,IND8=?,IND9=?,TP1=?,TP2=?,TP3=?,DRUG=?,DOSE=?,NOOFCYCLES=?,BSA=?,OTHERS=? WHERE REQNO=?"
 Set UPDINDICATIONQR = rdb.CreateQuery("UPDINDICATIONQR1", str)
 
 str = "INSERT INTO DISLIB.DBMACIND VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
 'STR = "INSERT INTO DISLIB.DBMACIND(REQNO,NEWCASE,POSTTRANS,DAY10MARR,PARTREATED,TSTEROIDS) VALUES(?,?,?,?,?,?)"
 Set INSDBMAINDQR = rdb.CreateQuery("INSDBMAINDQRBM", str)
 
 str = "INSERT INTO DISLIB.BLOODPRODU VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
 Set INSBLOODPRODUQR = rdb.CreateQuery("INSBLOODPRODUQRBP", str)
 
 str = "SELECT * FROM DISLIB.DBMACIND WHERE REQNO=?"
 Set SELDBMAINDQR = rdb.CreateQuery("SELDBMAINDQRBM", str)
 
 str = "SELECT * FROM DISLIB.DREQUISITN WHERE CASENO=? AND REQNO LIKE 'NZZSS%' "
 Set SELDREQNOQR = rdb.CreateQuery("SELDREQNOQR12", str)
 
 str = "SELECT * FROM DISLIB.BLOODPRODU WHERE REQNO=?"
 Set SELBLOODPRODUQR = rdb.CreateQuery("SELBLOODPRODUQRBP", str)

''added by Kanchan for Barcode
str = "SELECT * FROM PATOOLSLIB.MA106 WHERE DEPT_CODE=?"
Set SELDEPTDES_QR = rdb.CreateQuery("DEPTDESC1", str)
'''''''

str = "update RISLIB.RTIMELOG SET ADPREPDATE=?,ADPREPTIME=? WHERE REQNNO=?"
Set UPDRTIMEFORBC = rdb.CreateQuery("UPDRTIME11111", str)
 ''BP
 str = "SELECT * FROM DISLIB.DTMNO WHERE CASENO=?"
 Set SELDTMNO_QR = rdb.CreateQuery("SELDTMNO_STR", str)
 
 'STR = "select * from rislib.register where caseno=? AND (DAYS(REG_DATE)- DAYS(CURRENT DATE)) < 180"
 str = "SELECT * FROM RISLIB.REGISTER WHERE CASENO =? AND (DAYS(CURRENT DATE)- DAYS(REG_DATE)) <= 180"
 Set NOCHARG_QR = rdb.CreateQuery("NOCHARPRE", str)
 NOCHARG_FLG = False
  pathologyFLAG = True
  
  ''' sandhya 1/8/2008 protocol case
    
    str = "SELECT * FROM PATOOLSLIB.PROTOCASE WHERE CASE_NO = ?"
    Set SELprcase_QR = rdb.CreateQuery("SELPRCASE_QR", str)
    
    str = "SELECT * FROM PATOOLSLIB.PROTSERVIC WHERE PRJNO = ?"
    Set PROTSERV_QR = rdb.CreateQuery("PROTSERV_AQR", str)

'' end of protocol patient
  
''        Else
''            datetime
''            txt_date.Text = DATE11
''            txt_time.Text = time1
''        End If
 ' **** PROVISIONAL DIAGNOSIS COMBO BOX FILLING

    If SELMA10121FLAG = False Then
        SELMA10121
        Set SELMA10121_RS = SELMA10121_QR.OpenResultset(1, 3)
        SELMA10121FLAG = True
    Else
        SELMA10121
         Set SELMA10121_RS = SELMA10121_QR.OpenResultset(1, 3)
        SELMA10121_RS.MoreResults
        SELMA10121_RS.Requery
    End If

'    If SELMA10121_RS.RowCount > 0 Then
'        cmb_prodis.clear
'        While Not SELMA10121_RS.EOF
'            cmb_prodis.AddItem Trim(SELMA10121_RS("PRVDIAGDES"))
'            SELMA10121_RS.MoveNext
'        Wend
'    End If
    ' **** PROVISIONAL DIAGNOSIS COMBO BOX FILLING
    ' **** DEPARTMENT NAME COMBO BOX FILLING
    If SELMA106FLAG = False Then
        SELMA106
        Set SELMA106_RS = SELMA106_QR.OpenResultset(1, 3)
        SELMA106FLAG = True
    Else
        SELMA106
        Set SELMA106_RS = SELMA106_QR.OpenResultset(1, 3)
        SELMA106_RS.MoreResults
        SELMA106_RS.Requery
    End If
    If SELMA106_RS.RowCount > 0 Then
        cmb_dept.clear
        While Not SELMA106_RS.EOF
        cmb_dept.AddItem Trim(SELMA106_RS("DEPT_NAME")) & "$" & SELMA106_RS("DEPT_CODE")
        SELMA106_RS.MoveNext
        Wend
    End If
    ' **** DEPARTMENT NAME COMBO BOX FILLING
    ' **** REFERRING DOCTOR AND REFERRED TO DOCTOR FILLING
''            If SELEMPBASINFOFLAG = False Then
''                Set SELEMPBASINFO_RS = SELEMPBASINFO_QR.OpenResultset(1, 3)
''                SELEMPBASINFOFLAG = True
''            Else
''                SELEMPBASINFO_RS.MoreResults
''                SELEMPBASINFO_RS.Requery
''            End If
''            If SELEMPBASINFO_RS.RowCount > 0 Then
''                cmb_refdoc.Clear
''                cmb_reftodoc.Clear
''                While Not SELEMPBASINFO_RS.EOF
''                cmb_refdoc.AddItem Trim(SELEMPBASINFO_RS("EMPTITLE"))
''                cmb_reftodoc.AddItem Trim(SELEMPBASINFO_RS("EMPTITLE"))
''                SELEMPBASINFO_RS.MoveNext
''                Wend
''            End If
        If SELMA1031FLAG = False Then
            Set SELMA1031_RS = SELMA1031_QR.OpenResultset(1, 1)
            SELMA1031FLAG = True
        Else
            Set SELMA1031_RS = SELMA1031_QR.OpenResultset(1, 1)
            SELMA1031_RS.MoreResults
            SELMA1031_RS.Requery
        End If
        If SELMA1031_RS.RowCount > 0 Then
            cmb_refdoc.clear
            cmb_reftodoc.clear
            While Not SELMA1031_RS.EOF
                cmb_refdoc.AddItem Trim(SELMA1031_RS("UD_NAME"))
                cmb_reftodoc.AddItem Trim(SELMA1031_RS("UD_NAME"))
                SELMA1031_RS.MoveNext
            Wend
            
        End If
    ' **** REFERRING DOCTOR AND REFERRED TO DOCTOR FILLING
End If

'************** Shraddha Pandit 02/03/09 ***************
If ipadd = "150.150.200.62" Or ipadd = "150.150.15.51" Then
    lbl_icureq.Visible = True
End If
If ipadd = "10.3.1.56" Then
    lbl_rwd.Visible = True
End If
'*******************************************************



'''venkat*******
'    str12 = "select * from dislib.sitedup where sitedesc = ?"
'    Set sitedupqr = rdb.CreateQuery("site100", str12)
'     strins = "insert into dislib.sitedup(sitecd,sitedesc) values(?,?)"
'     Set siteinsertqr = rdb.CreateQuery("siteInsertdup1", strins)


'CMB_ICU1.AddItem "NO"
btnenabl ("INI")
FLDENABL ("INI")

'added by milind on 17/04/06 to get a link from ot module
    
    VAR = COMMAND()
    If Len(VAR) > 0 Then
        UID = "A" + "DISCLK"
        If Mid(Trim(UID), 2) = "DISCLK" Then
            opt_patient.Value = True
            cmd_new_Click
            Msk_caseno.Text = VAR
            msk_caseno_LostFocus
        End If
    End If

'upto here
NEWFLAG_BCD = False
VIEWFLAG_BCD = False

Exit Sub
MSG:
MsgBox Err.Description
End Sub

Public Sub btnenabl(cap As String)

If cap = "NEW" Then
    cmd_save.Enabled = True
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_new.Enabled = False
    cmd_modify.Enabled = False
    cmd_print.Enabled = False
    cmd_repeat.Enabled = False
    cmd_view.Enabled = False
    CMB_MEMO.Enabled = False
ElseIf cap = "VIEW" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_modify.Enabled = True
    cmd_new.Enabled = False
    cmd_save.Enabled = False
    cmd_print.Enabled = True
    cmd_repeat.Enabled = False
    cmd_view.Enabled = False
    CMB_MEMO.Enabled = True
ElseIf cap = "SAVE" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_modify.Enabled = False
    cmd_new.Enabled = True
    cmd_save.Enabled = False
    cmd_print.Enabled = True
    Check1.Value = 0
    If Trim(txt_dept.Text) = "FEC" Then
        cmd_repeat.Enabled = False
    Else
        cmd_repeat.Enabled = True
    End If
    'cmd_repeat.Enabled = True
    cmd_view.Enabled = True
    If (Trim(txt_catgory.Text) = "NC") Or (opt_staff.Value = True) Then
        CMB_MEMO.Enabled = False
    Else
        CMB_MEMO.Enabled = True
    End If
    
    
    
ElseIf cap = "INI" Then
    cmd_cancel.Enabled = False
    cmd_back.Enabled = True
    cmd_modify.Enabled = False
    cmd_new.Enabled = True
    cmd_save.Enabled = False
    cmd_print.Enabled = False
    cmd_repeat.Enabled = False
    cmd_view.Enabled = True
    CMB_MEMO.Enabled = False
ElseIf cap = "REPEAT" Then
    cmd_save.Enabled = True
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_new.Enabled = False
    cmd_modify.Enabled = False
    cmd_print.Enabled = False
    cmd_repeat.Enabled = False
    cmd_view.Enabled = False
    CMB_MEMO.Enabled = True
ElseIf cap = "PRINT" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_modify.Enabled = False
    cmd_new.Enabled = True
    cmd_save.Enabled = False
    cmd_print.Enabled = True
    cmd_repeat.Enabled = True
    cmd_view.Enabled = True
    If Trim(txt_dept.Text) = "FEC" Then
        cmd_repeat.Enabled = False
    Else
        cmd_repeat.Enabled = True
    End If
    If (Trim(txt_catgory.Text) = "NC" Or (opt_staff.Value = True)) Then
        CMB_MEMO.Enabled = False
    Else
        CMB_MEMO.Enabled = True
    End If
   
End If
'' SANDHYA PROTOCOL
Label12.Enabled = False
Label12.Visible = False
CMBIRB.Enabled = False
CMBIRB.Visible = False
'' END OF PROTOCOL
End Sub
Public Sub CY()

    
    
    '**** CYTOLOGY
        
  If selreqnno_rs.RowCount > 0 Then
     cytogeneticreq.cmb_cyspreinv.clear
     selreqnno_rs.MoveFirst
     While Not selreqnno_rs.EOF
       selreqmoda_qr(0) = selreqnno_rs(0)
       Set selreqmoda_RS = selreqmoda_qr.OpenResultset(2, 3)
       cytogeneticreq.cmb_cyspreinv.AddItem selreqnno_rs(0) & "   " & selreqmoda_RS(0)
     selreqnno_rs.MoveNext
     Wend
  End If
    
        If Trim(pathology.txt_modcode.Text) = "CC" Or Trim(pathology.txt_modcode.Text) = "FI" Then
               If SELRINVESTCY_FLAG = False Then
                SELRINVESTCY_QR(0) = Trim(pathology.txt_modcode.Text)
                Set SELRINVESTCY_RS = SELRINVESTCY_QR.OpenResultset(1, 3)
                SELRINVESTCY_FLAG = True
            Else
                SELRINVESTCY_QR(0) = Trim(pathology.txt_modcode.Text)
                SELRINVESTCY_RS.MoreResults
                SELRINVESTCY_RS.Requery
            End If
            If SELRINVESTCY_RS.RowCount > 0 Then
                cmb_cyprocedure.clear
                cmb_cyprocode.clear
                While Not SELRINVESTCY_RS.EOF
                    cmb_cyprocode.AddItem Trim(SELRINVESTCY_RS("ser_code"))
                    'cmb_procedurecode.Text = Trim(SELRINVEST1_RS("ser_code"))
                    If SELMA101CG_FLAG = False Then
                        SELMA101CY_QR(0) = Trim(SELRINVESTCY_RS("SER_CODE"))
                        Set SELMA101CY_RS = SELMA101CY_QR.OpenResultset(1, 3)
                        SELMA101CY_FLAG = True
                     Else
                        SELMA101CY_QR(0) = SELRINVESTCY_RS("SER_CODE")
                        SELMA101CY_RS.MoreResults
                        SELMA101CY_RS.Requery
                    End If
                    If SELMA101CY_RS.RowCount > 0 Then
                        While Not SELMA101CY_RS.EOF
                        cmb_cyprocedure.AddItem SELMA101CY_RS("SER_DET")
                        'cmb_procedure.Text SELMA101_RS1("SER_DET")
                        SELMA101CY_RS.MoveNext
                        Wend
                    End If
                    SELRINVESTCY_RS.MoveNext
                Wend
            End If
        End If
 

 Dim I As Integer
   I = InStr(cmb_modality.Text, "$")
   txt_modcode.Text = Trim(Mid(cmb_modality.Text, I + 1, 2))
   cytogeneticreq.show
   Me.Hide
End Sub

Public Sub CYASSIGN()


  '*** INSERTING INT0 RREQUISITION

  
  
INSRREQUISITN_QR(0) = req_chglen(Trim(txt_bcreqno.Text))
INSRREQUISITN_QR(1) = GET_REVDATE(CStr(DATE11))
INSRREQUISITN_QR(2) = Trim(txt_time.Text)
INSRREQUISITN_QR(3) = Trim(Msk_caseno.FormattedText)
INSRREQUISITN_QR(4) = Trim(txt_dept.Text)
INSRREQUISITN_QR(5) = Trim(txt_modcode.Text)
INSRREQUISITN_QR(6) = Trim(txt_status.Text)
INSRREQUISITN_QR(7) = Trim(txt_catgory.Text)
INSRREQUISITN_QR(8) = Trim(cmb_cyprocode.Text)
INSRREQUISITN_QR(9) = UCase(Trim(cmb_refdoc.Text))

'Added by Milind on 27-03-2006 for inserting date for the samples collected at outside
    If Check1.Value = 1 Then
        Dim Dt As String
        Dt = "18/02/1980"
        INSRREQUISITN_QR(15) = Format(CDate(Dt), "dd/mm/yyyy")
    Else
        INSRREQUISITN_QR(15) = Null
    End If
'upto here


    If SELRREQUISITNFLAG = False Then
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
        SELRREQUISITNFLAG = True
    Else
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        SELRREQUISITN_RS.MoreResults
        SELRREQUISITN_RS.Requery
    End If
    If SELRREQUISITN_RS.RowCount > 0 Then
        INSRREQUISITN_QR(10) = "YES"
    Else
        INSRREQUISITN_QR(10) = "NO"
    End If
    
    If pathology.opt_patient.Value = True Then
        NO_CHARGE
      If No_Chargeflag = True Then
            INSRREQUISITN_QR(11) = "NOCHARGE"
            INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
            INSRREQUISITN_QR(13) = "P"
            No_Chargeflag = False
      Else
            INSRREQUISITN_QR(11) = Null
            INSRREQUISITN_QR(12) = Null
            INSRREQUISITN_QR(13) = Null
        
      End If
    End If
    
    If pathology.opt_staff.Value = True Then
            INSRREQUISITN_QR(11) = "NOCHARGE"
            INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
            INSRREQUISITN_QR(13) = "P"
    End If

    INSRREQUISITN_QR(14) = 0
    INSRREQUISITN_QR.Execute
    
    
    INSCYTOGCLDTACY_QR(0) = req_chglen(Trim(txt_bcreqno.Text))
    INSCYTOGCLDTACY_QR(1) = Trim(Msk_caseno.FormattedText)
    INSCYTOGCLDTACY_QR(2) = cytogeneticreq.chk_at_dig.Value
    INSCYTOGCLDTACY_QR(3) = cytogeneticreq.chk_pr_treat.Value
    INSCYTOGCLDTACY_QR(4) = cytogeneticreq.chk_ps_indu.Value
    INSCYTOGCLDTACY_QR(5) = cytogeneticreq.chk_relapse.Value
    INSCYTOGCLDTACY_QR(6) = cytogeneticreq.chk_remission.Value
    INSCYTOGCLDTACY_QR(7) = Trim(cytogeneticreq.txt_followup.Text)
    INSCYTOGCLDTACY_QR(8) = cytogeneticreq.chk_postbmt.Value
    INSCYTOGCLDTACY_QR(9) = Trim(cytogeneticreq.cmb_sex.Text)
    INSCYTOGCLDTACY_QR(10) = Trim(cytogeneticreq.TXT_OTHERS.Text)
    INSCYTOGCLDTACY_QR(11) = Trim(StrCgdiag1)
    INSCYTOGCLDTACY_QR(12) = Trim(StrCgdiag2)
    INSCYTOGCLDTACY_QR(13) = Val(Trim(cytogeneticreq.txt_phone.Text))
    INSCYTOGCLDTACY_QR(14) = Val(Trim(cytogeneticreq.txt_fax.Text))
    INSCYTOGCLDTACY_QR(15) = ""
    INSCYTOGCLDTACY_QR(16) = ""
    INSCYTOGCLDTACY_QR.Execute

        'inserting into drequisitn
    
        INSDREQUISITN_QR(0) = req_chglen(Trim(txt_bcreqno.Text))
        INSDREQUISITN_QR(1) = " "
        'INSDREQUISITN_QR(1) = Trim(txt_sppathno.Text)
        INSDREQUISITN_QR(2) = Trim(Msk_caseno.FormattedText)
        INSDREQUISITN_QR(3) = Null
        INSDREQUISITN_QR(4) = Null
        INSDREQUISITN_QR(5) = Null
        INSDREQUISITN_QR(6) = Null
        INSDREQUISITN_QR(7) = Null
        INSDREQUISITN_QR(8) = Null
        INSDREQUISITN_QR(9) = Null
        INSDREQUISITN_QR(10) = Null
        INSDREQUISITN_QR(11) = Null
        INSDREQUISITN_QR(12) = Null
        INSDREQUISITN_QR(13) = Null
        INSDREQUISITN_QR(14) = Trim(cmb_refdoc.Text)
        INSDREQUISITN_QR(15) = Trim(cmb_reftodoc.Text)
        INSDREQUISITN_QR(16) = Trim(cmb_cyprocedure.Text)
        
        If Trim(pathology.txt_status.Text) = "IP" Then
            INSDREQUISITN_QR(17) = Trim(txt_ward.Text)
            INSDREQUISITN_QR(18) = Trim(txt_bedno.Text)
        Else
            INSDREQUISITN_QR(17) = Null
            INSDREQUISITN_QR(18) = Null
        End If
        INSDREQUISITN_QR(19) = " "
        INSDREQUISITN_QR(20) = Null
        INSDREQUISITN_QR(21) = Trim(ICU)
        INSDREQUISITN_QR.Execute
        
        '**** INSERTING INTO DREQUISITION
        
        If SELTA1012FLAG = False Then
            selta1012_qr(0) = Msk_caseno.FormattedText
            Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
            SELTA1012FLAG = True
        Else
             selta1012_qr(0) = Msk_caseno.FormattedText
             selta1012_rs.MoreResults
             selta1012_rs.Requery
        End If
        
        If selta1012_rs.RowCount > 0 Then
            UPDTA1012_QR(0) = Trim(txt_indctinvstg.Text)
            UPDTA1012_QR(1) = Msk_caseno.FormattedText
            UPDTA1012_QR.Execute
        Else
            insta1012_qr(0) = Msk_caseno.FormattedText
            insta1012_qr(1) = Trim(txt_indctinvstg.Text)
            insta1012_qr.Execute
        End If
'cmb_cyprocedure.clear
'cmb_cyprocode.clear

Unload cytogeneticreq

Data1.UpdateRecord
Data1.Recordset.Bookmark = Data1.Recordset.LastModified
   
Exit Sub
MSG:
End Sub

Public Sub cyview()
Dim K As Integer
 If SELCYTOGCLDTACY_FLG = False Then
   SELCYTOGCLDTACY_QR(0) = req_chglen(cmb_reqno.Text)
   Set SELCYTOGCLDTACY_RS = SELCYTOGCLDTACY_QR.OpenResultset(2, 3)
   SELCYTOGCLDTACY_FLG = True
 Else
   SELCYTOGCLDTACY_RS.MoreResults
   SELCYTOGCLDTACY_RS.Requery
 End If
 If SELCYTOGCLDTACY_RS.RowCount > 0 Then
    cytogeneticreq.show
    cytogeneticreq.chk_at_dig.Value = SELCYTOGCLDTACY_RS!AT_DIaG
    cytogeneticreq.chk_pr_treat.Value = SELCYTOGCLDTACY_RS!pr_treat
    cytogeneticreq.chk_ps_indu.Value = SELCYTOGCLDTACY_RS!ps_indu
    cytogeneticreq.chk_relapse.Value = SELCYTOGCLDTACY_RS!relapse
    cytogeneticreq.chk_remission.Value = SELCYTOGCLDTACY_RS!remission
    cytogeneticreq.txt_followup.Text = SELCYTOGCLDTACY_RS!FOLLUP
    If cytogeneticreq.txt_followup = "FOLLOW-UP" Then
        cytogeneticreq.Option1.Value = True
    '    cytogeneticreq.Option2.Value = False
    Else
        cytogeneticreq.Option2.Value = True
    End If
    
    cytogeneticreq.chk_postbmt.Value = SELCYTOGCLDTACY_RS!postbmt
    cytogeneticreq.cmb_sex.Text = SELCYTOGCLDTACY_RS!B_SEX
    cytogeneticreq.TXT_OTHERS.Text = SELCYTOGCLDTACY_RS!SP_TYPE
    StrCgdiag1 = SELCYTOGCLDTACY_RS!CGDIAG1
    StrCgdiag2 = SELCYTOGCLDTACY_RS!CGDIAG2
    If SELCYTOGCLDTACY_RS!SP_TYPE = "Bone Marrow Aspiration" Then
      cytogeneticreq.opt_BMA.Value = 1
    'End If
    ElseIf SELCYTOGCLDTACY_RS!SP_TYPE = "Peripheral Blood" Then
      cytogeneticreq.opt_PBlood.Value = 1
    'End If
    ElseIf SELCYTOGCLDTACY_RS!SP_TYPE = "Lymph Node" Then
      cytogeneticreq.opt_lymbnode.Value = 1
    Else
      cytogeneticreq.opt_option.Value = 1
    
       
    End If
    K = InStr(1, StrCgdiag1, "AL")
     If K > 0 Then
       cytogeneticreq.chk_al.Value = 1
     End If
    K = InStr(1, StrCgdiag1, "AML")
     If K > 0 Then
       cytogeneticreq.chk_aml.Value = 1
     End If
     K = InStr(1, StrCgdiag1, "ALL")
     If K > 0 Then
       cytogeneticreq.chk_all.Value = 1
     End If
     K = InStr(1, StrCgdiag1, "CML")
     If K > 0 Then
       cytogeneticreq.chk_cml.Value = 1
     End If
     K = InStr(1, StrCgdiag1, "NHL")
     If K > 0 Then
       cytogeneticreq.chk_nhl.Value = 1
     End If
     K = InStr(1, StrCgdiag1, "LYMPHOMA")
     If K > 0 Then
       cytogeneticreq.chk_lymphoma.Value = 1
     End If
     K = InStr(1, StrCgdiag1, "FANCONI'S ANEMIA")
     If K > 0 Then
       cytogeneticreq.chk_fanane.Value = 1
     End If
     K = InStr(1, StrCgdiag1, "APLASTIC ANEMIA")
     If K > 0 Then
       cytogeneticreq.chk_aplane.Value = 1
     End If
     K = InStr(1, StrCgdiag1, "MDS")
     If K > 0 Then
       cytogeneticreq.chk_mds.Value = 1
     End If
     K = InStr(1, StrCgdiag2, "AML-M0")
     If K + 0 Then
          cytogeneticreq.chk_amlm0.Value = 1
     End If
     K = InStr(1, StrCgdiag2, "AML-M1")
     If K + 0 Then
          cytogeneticreq.chk_amlm1.Value = 1
     End If
     K = InStr(1, StrCgdiag2, "AML-M2")
     If K + 0 Then
          cytogeneticreq.chk_amlm2.Value = 1
     End If
     K = InStr(1, StrCgdiag2, "AML-M3")
     If K + 0 Then
          cytogeneticreq.chk_amlm3.Value = 1
     End If
     K = InStr(1, StrCgdiag2, "AML-M4")
     If K + 0 Then
          cytogeneticreq.chk_amlm4.Value = 1
     End If
     K = InStr(1, StrCgdiag2, "AML-M4/E0")
     If K + 0 Then
          cytogeneticreq.chk_amlm4e0.Value = 1
     End If
     K = InStr(1, StrCgdiag2, "AML-M5a")
     If K + 0 Then
          cytogeneticreq.chk_m5a.Value = 1
     End If
     K = InStr(1, StrCgdiag2, "AML-M5b")
     If K + 0 Then
          cytogeneticreq.chk_m5b.Value = 1
     End If
     K = InStr(1, StrCgdiag2, "AML-M6")
     If K + 0 Then
          cytogeneticreq.chk_aml6.Value = 1
          
     End If
     K = InStr(1, StrCgdiag2, "AML-M7")
     If K + 0 Then
          cytogeneticreq.chk_amlm7.Value = 1
          
     End If
     K = InStr(1, StrCgdiag2, "BIPHENOTYPIC LEUKEMIA")
     If K + 0 Then
          cytogeneticreq.chk_bioleu.Value = 1
     End If
     K = InStr(1, StrCgdiag2, "Pre B-ALL")
     If K + 0 Then
          cytogeneticreq.chk_preball.Value = 1
     End If
     K = InStr(1, StrCgdiag2, "T-ALL")
     If K + 0 Then
          cytogeneticreq.chk_tall.Value = 1
     End If
     K = InStr(1, StrCgdiag2, "CML-CP")
     If K + 0 Then
          cytogeneticreq.chk_cmlcp.Value = 1
     End If
     K = InStr(1, StrCgdiag2, "CML-BC")
     If K + 0 Then
          cytogeneticreq.chk_cmlbc.Value = 1
     End If
     K = InStr(1, StrCgdiag2, "CML-AP")
     If K + 0 Then
          cytogeneticreq.chk_cmlap.Value = 1
     End If
     K = InStr(1, StrCgdiag2, "Burkitt")
     If K + 0 Then
          cytogeneticreq.chk_burkitt.Value = 1
     End If
     K = InStr(1, StrCgdiag2, "CLL")
      If K + 0 Then
          cytogeneticreq.chk_cll.Value = 1
     End If
     
     
    cytogeneticreq.txt_fax = SELCYTOGCLDTACY_RS!FAX
    cytogeneticreq.txt_phone.Text = SELCYTOGCLDTACY_RS!PHONE
 End If
 
 
End Sub

Public Sub CYPRINT()
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPCY.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(cmb_reqno) & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                
                If txt_modcode.Text = "CC" Or txt_modcode.Text = "FI" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR CYTOGENETICS" & " '"
                End If
                    CrystalReport1.Formulas(7) = "ENTERBY='" & Trim(enter_by1) & "'"
                    CrystalReport1.Formulas(8) = "prov='" & Trim(txt_indctinvstg.Text) & "'"
                    CrystalReport1.Formulas(9) = "details='" & Trim(detail) & "'"
                
End Sub

Private Sub Frame1_DragDrop(Source As CONTROL, X As Single, Y As Single)

End Sub

Private Sub Form_Unload(cancel As Integer)
'Dim RS As rdoResultset
'Dim RQ As rdoQuery
'For Each RS In rdb.rdoResultsets
'    RS.Close
'Next
'
'For Each RQ In rdb.rdoQueries
'    RQ.Close
'Next
End Sub

Private Sub list_proc_Click()
I = I1 - 1
While I >= 0

    If list_proc.Selected(I) Then
        List_code.Selected(I) = True
    
         
         If List_code.Text = "G253" Then
            
            list_proc.RemoveItem (g119cd)
            List_code.RemoveItem (g119cd)
            list_proc.RemoveItem (g106cd)
            List_code.RemoveItem (g106cd)

         End If
         If List_code.Text = "G106" Or List_code.Text = "G119" Then
            'MsgBox ("hI"), vbOKOnly
            'list_code.Index
            list_proc.RemoveItem (g253cd)
            List_code.RemoveItem (g253cd)
        End If
            
       
    ElseIf List_code.Selected(I) Then
        List_code.Selected(I) = False
    End If
    I = I - 1
Wend


End Sub

Private Sub List_code_Click()
I = I1 - 1
While I >= 0
    If List_code.Selected(I) Then
        list_proc.Selected(I) = True
    ElseIf list_proc.Selected(I) Then
        list_proc.Selected(I) = False
    End If
    I = I - 1
Wend

End Sub
'********** Added by shraddha on 02/07/09 ***********
Private Sub list_proc_LostFocus()
'SendKeys "^{f4}"
cmb_specimen.SetFocus
End Sub
'****************************************************


Private Sub msk_caseno_LostFocus()
On Error GoTo MSG:
If opt_patient.Value = False And opt_staff.Value = False Then
    Exit Sub
End If
If Trim(Msk_caseno.Text) = "" And opt_patient.Value = True Then
MsgBox "Case No. is Empty"
opt_patient.SetFocus
'btnenabl ("INI")
Exit Sub
End If
If opt_patient.Value = True Then
    CASENO = padzero(5, "C", Msk_caseno.Text, 3)
    Msk_caseno.Text = UCase(CASENO)
    If SELTA101FALG = False Then
        'selta101
        SELTA101_QR(0) = Msk_caseno.FormattedText
        Set SELTA101_RS = SELTA101_QR.OpenResultset(2, 3)
        SELTA101FALG = True
    Else
        'selta101
        SELTA101_QR(0) = Msk_caseno.FormattedText
        Set SELTA101_RS = SELTA101_QR.OpenResultset(2, 3)
        SELTA101_RS.MoreResults
        SELTA101_RS.Requery
    End If
    RETRIVE
    If Trim(txt_catgory.Text) = "NC" Then
        CMB_MEMO.Enabled = False
    End If
    
    If viewflag = True Then
        cmb_reqno.Enabled = True
        cmb_reqno.SetFocus
    End If
End If
If opt_staff.Value = True Then
''    While Len(Msk_caseno.Text) < 8
''     Msk_caseno = "0" & Msk_caseno
''    Wend
    
    If SELEMPBASINFOFLAG = False Then
        SELEMPBASINFO_QR(0) = Trim(Msk_caseno.Text)
        Set SELEMPBASINFO_RS = SELEMPBASINFO_QR.OpenResultset(1, 3)
        SELEMPBASINFOFLAG = True
    Else
        SELEMPBASINFO_QR(0) = Trim(Msk_caseno.Text)
        SELEMPBASINFO_RS.MoreResults
        SELEMPBASINFO_RS.Requery
    End If
    
    'added by milind on 22-03-2007 for disabling requisitions for staff
    
    If SELELEMPFLAG = False Then
        SELELEMP_QR(0) = Trim(Msk_caseno.Text)
        Set SELELEMP_RS = SELELEMP_QR.OpenResultset(1, 3)
        SELELEMPFLAG = True
    Else
        SELELEMP_QR(0) = Trim(Msk_caseno.Text)
        SELELEMP_RS.MoreResults
        SELELEMP_RS.Requery
    End If
    
    If SELELEMP_RS.RowCount > 0 Then
        'If Trim(SELELEMP_RS("CESSYN")) <> "4" Then
         If Trim(SELELEMP_RS("CESSYN")) <> "4" And Trim(SELELEMP_RS("CESSYN")) <> "9" Then
            
                MsgBox "THIS EMPLOYEE IS NOT CURRENTLY WORKING WITH TATA HOSPITAL !!!!!", vbInformation
                'Msk_caseno.Text = ""
                Exit Sub
            
        End If
    End If
    
    'upto here
    
    RETRIVE
    CMB_MEMO.Enabled = False
End If

If cmb_refdoc.Text = "NO UNIT" Then
    MsgBox "Please Enter Proper Unit", vbExclamation, "DIS"
    cmb_refdoc.Enabled = False
    'cmd_save.Enabled = False
    Exit Sub
End If
'' SANDHYA PROTOCOL PATIENT
If prflg = True Then
  CMBIRB.SetFocus
  MsgBox ("PLEASE ENTER PROJECT CODE FOR THIS PROTOCOL PATIENT")
End If
'' END OF PROTOCOL

Exit Sub
MSG:
MsgBox Err.Description
End Sub
Public Sub RETRIVE()
Dim YEAR1 As Integer
On Error GoTo MSG:
If opt_patient.Value = True Then
If SELTA101_RS.RowCount > 0 Then
    If NEWFLAG = True Then
           'newflag = False
           txt_name.Text = Trim(SELTA101_RS("TAG")) & " " & Trim(SELTA101_RS("FIRST_NAME")) & " " & Trim(SELTA101_RS("MID_NAME")) & " " & Trim(SELTA101_RS("LAST_NAME"))
           txt_sex = Trim(SELTA101_RS("SEX"))
           PATYPE = Trim(SELTA101_RS("PAT_TYPE"))
           txt_Age = DateDiff("YYYY", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
'               If Len(Trim(SELTA101_RS("BIRTH_DATE"))) <> 0 Then
'                   YEAR1 = DateDiff("m", SELTA101_RS("BIRTH_DATE"), DATE11)
'                   Dim N As Integer
'                   N = YEAR1 / 12
'                   txt_Age.Text = N
'               End If
'
'               If Trim(txt_Age.Text) = 0 Then
'                  Dim age As String
'                  age = DateDiff("m", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
'                  age = "." & (age Mod 12)
'                     If age = ".0" Then
'                        age = DateDiff("d", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
'                        age = ".0" & (age) ' Mod 30)
'                     End If
'                  txt_Age.Text = Trim(age)
'               End If
               
           
'           If selprvdgs = False Then
'                selprvdgs_qr(0) = msk_caseno.FormattedText
'                Set selprvdgs_rs = selprvdgs_qr.OpenResultset(1, 3)
'                selprvdgs = True
'           Else
'                selprvdgs_qr(0) = msk_caseno.FormattedText
'                selprvdgs_rs.MoreResults
'                selprvdgs_rs.Requery
'           End If
'                txt_indctinvstg.Text = Trim(selprvdgs_rs("prvdiagdes"))
           
           txt_catgory = Trim(SELTA101_RS("PAT_CATG"))
           If selma104flag = False Then
               selma104_qr(0) = Msk_caseno.FormattedText
               Set selma104_rs = selma104_qr.OpenResultset(1, 3)
               selma104flag = True
           Else
               selma104_qr(0) = Msk_caseno.FormattedText
               selma104_rs.MoreResults
               selma104_rs.Requery
           End If
           If selma104_rs.RowCount > 0 Then
               txt_status = "IP"
               txt_bedno.Visible = True
               txt_ward.Visible = True
               lbl_bedno.Visible = True
               lbl_ward.Visible = True
               txt_bedno.Text = Trim(selma104_rs("BED_NO"))
               txt_ward.Text = Trim(selma104_rs("WARD_NAME"))
           Else
               txt_status = "OP"
               txt_bedno.Visible = False
               txt_ward.Visible = False
               lbl_bedno.Visible = False
               lbl_ward.Visible = False
               
           End If
           txt_pridocser.Text = Trim(SELTA101_RS("UD_NAME"))
                
           'cmb_refdoc.Text = Trim(SELTA101_RS("UD_NAME"))
           If SELTA1011FLAG = False Then
               SELTA1011_QR(0) = Msk_caseno.FormattedText
               Set SELTA1011_RS = SELTA1011_QR.OpenResultset(2, 3)
               SELTA1011FLAG = True
           Else
               SELTA1011_QR(0) = Msk_caseno.FormattedText
               SELTA1011_RS.MoreResults
               SELTA1011_RS.Requery
           End If
           If SELTA1011_RS.RowCount > 0 Then
                If Left(Msk_caseno.FormattedText, 2) = "RF" Or Left(Msk_caseno.FormattedText, 2) = "RG" Then
                    
                    cmb_refdoc = SELTA1011_RS("REF_BY_DOC")
                    cmb_reftodoc.Visible = True
                    lbl_reftodoc.Visible = True
                Else
                    cmb_refdoc.Text = Trim(SELTA101_RS("UD_NAME"))
                    cmb_reftodoc.Visible = False
                    lbl_reftodoc.Visible = False
               End If
           Else
                If Left(Msk_caseno.FormattedText, 2) = "RF" Or Left(Msk_caseno.FormattedText, 2) = "RG" Then
                    
                    cmb_refdoc = ""
                    cmb_reftodoc.Visible = True
                    lbl_reftodoc.Visible = True
                Else
                    cmb_refdoc.Text = Trim(SELTA101_RS("UD_NAME"))
                    cmb_reftodoc.Visible = False
                    lbl_reftodoc.Visible = False
               End If
               
           End If
           If SELCLINHISTFLAG = False Then
               SELCLINHIST_QR(0) = Msk_caseno.FormattedText
               Set SELCLINHIST_RS = SELCLINHIST_QR.OpenResultset(1, 3)
               SELCLINHISTFLAG = True
           Else
               SELCLINHIST_QR(0) = Msk_caseno.FormattedText
               SELCLINHIST_RS.MoreResults
               SELCLINHIST_RS.Requery
           End If
           If SELCLINHIST_RS.RowCount > 0 Then
                SELCLINHIST_RS.MoveLast
                If Trim(SELCLINHIST_RS("CLINHIST")) = "" Then
                    txt_clinhist.Text = "NYD"
               Else
                txt_clinhist.Text = Trim(SELCLINHIST_RS("CLINHIST"))
               End If
           Else
               txt_clinhist.Text = "NYD"
           End If
           If txt_status = "IP" And txt_catgory = "NC" Then
                INNC = True
           End If
        End If
        If viewflag = True Then
           
            If SELRREQUISITNFLAG = False Then
                SELRREQUISITN_QR(0) = Msk_caseno.FormattedText
                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
                SELRREQUISITNFLAG = True
             Else
                SELRREQUISITN_QR(0) = Msk_caseno.FormattedText
                SELRREQUISITN_RS.MoreResults
                SELRREQUISITN_RS.Requery
            
            End If
            If SELRREQUISITN_RS.RowCount > 0 Then
                cmb_reqno.clear
                While Not SELRREQUISITN_RS.EOF
                    If Dept_cd1 = "ALL" Then
                        cmb_reqno.AddItem req_chglen(SELRREQUISITN_RS("REQNNO"))
                        SELRREQUISITN_RS.MoveNext
                    Else
                       If Mid(SELRREQUISITN_RS("REQNNO"), 1, 3) = Dept_cd1 Then
                            cmb_reqno.AddItem req_chglen(SELRREQUISITN_RS("REQNNO"))
                        End If
                        SELRREQUISITN_RS.MoveNext
                    End If
                Wend
             Else
                MsgBox " Case No. is not Registered"
                Exit Sub
            End If
        End If
        If txt_status = "IP" And txt_catgory = "NC" Then
                INNC = True
        End If
        
        Else
        MsgBox " Case No. is not Registered.", vbInformation
        btnenabl ("INI")
        Exit Sub

    End If
    
'' SANDHYA 1.8.2008 PROTOCOL PATIENT
SELprcase_QR(0) = Msk_caseno.FormattedText
Set SELPRCASE_RS = SELprcase_QR.OpenResultset(2, 3)

If SELPRCASE_RS.RowCount > 0 Then
   Label12.Enabled = True
   Label12.Visible = True
   CMBIRB.Enabled = True
   CMBIRB.Visible = True
   CMBIRB.clear
   While Not SELPRCASE_RS.EOF
       CMBIRB.AddItem Trim(SELPRCASE_RS("IRBNO"))
       SELPRCASE_RS.MoveNext
   Wend
   prflg = True
   
   'CMBIRB.SetFocus
Else
   prflg = False
   Label12.Enabled = False
   Label12.Visible = False
   CMBIRB.Enabled = False
   CMBIRB.Visible = False
End If

'' END OF PROTOCOL PATIENT
    
End If
If opt_staff.Value = True Then
    
        If NEWFLAG = True Then
                'newflag = False
           
                If SELEMPBASINFO_RS.RowCount > 0 Then
                     txt_name.Text = Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME"))
                     txt_sex = Trim(SELEMPBASINFO_RS("GENDER"))
                     txt_Age = DateDiff("YYYY", SELEMPBASINFO_RS("DOB"), DATE11)
                     txt_status = "ST"
                End If
            
               If SELEMPDPDLISTFLAG = False Then
                    SELEMPDPDLIST_QR(0) = Trim(Msk_caseno.Text)
                    Set SELEMPDPDLIST_RS = SELEMPDPDLIST_QR.OpenResultset(1, 3)
                    SELEMPDPDLISTFLAG = True
                Else
                    SELEMPDPDLIST_QR(0) = Trim(Msk_caseno.Text)
                    SELEMPDPDLIST_RS.MoreResults
                    SELEMPDPDLIST_RS.Requery
               End If
                If SELEMPDPDLIST_RS.RowCount > 0 Then
                     cmb_dependents.clear
                     While Not SELEMPDPDLIST_RS.EOF
                         cmb_dependents.AddItem Trim(SELEMPDPDLIST_RS("DPDNAME")) & "$" & Trim(SELEMPDPDLIST_RS("DPDREL"))
                     SELEMPDPDLIST_RS.MoveNext
                     Wend
                End If
           
           
           
                If SELCLINHISTFLAG = False Then
                    SELCLINHIST_QR(0) = Msk_caseno.FormattedText
                    Set SELCLINHIST_RS = SELCLINHIST_QR.OpenResultset(1, 3)
                    SELCLINHISTFLAG = True
                Else
                    SELCLINHIST_QR(0) = Msk_caseno.FormattedText
                    SELCLINHIST_RS.MoreResults
                    SELCLINHIST_RS.Requery
                End If
                If SELCLINHIST_RS.RowCount > 0 Then
                    SELCLINHIST_RS.MoveLast
                    txt_clinhist.Text = Trim(SELCLINHIST_RS("CLINHIST"))
                Else
                    txt_clinhist.Text = "NYD"
                End If
                If txt_status = "IP" And txt_catgory = "NC" Then
                     INNC = True
                End If
                
      End If
       
     If viewflag = True Then
               
                If SELRREQUISITNFLAG = False Then
                    SELRREQUISITN_QR(0) = Msk_caseno.FormattedText
                    Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
                    SELRREQUISITNFLAG = True
                 Else
                    SELRREQUISITN_QR(0) = Msk_caseno.FormattedText
                    SELRREQUISITN_RS.MoreResults
                    SELRREQUISITN_RS.Requery
                
                End If
                If SELRREQUISITN_RS.RowCount > 0 Then
                    cmb_reqno.clear
                    While Not SELRREQUISITN_RS.EOF
                        cmb_reqno.AddItem req_chglen(SELRREQUISITN_RS("REQNNO"))
                        SELRREQUISITN_RS.MoveNext
                    Wend
                Else
                    MsgBox " Case No. is not Registered"
                    Exit Sub
                End If
     
     End If
 
End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub
Private Sub CMBIRB_Click()
'' SANDHYA PROTOCOL PATIENT

If CMBIRB.Text = " " Then
CMBIRB.SetFocus
Else
  cmb_dept.SetFocus
End If
'' END OF PROTOCOL
End Sub
Private Sub CMBIRB_LostFocus()
'' SANDHYA PROTOCOL PATIENT
'CMBIRB.Text = MIRBNO
'MsgBox ("PLEASE ENTER PROJECT CODE FOR THIS PROTOCOL PATIENT")

MIRBNO = CMBIRB.Text
'cmb_dept.SetFocus
'' END OF PROTOCOL
End Sub


Public Sub SP()
On Error GoTo MSG:
If txt_modcode.Text = "SP" Then
    SPENABLE ("ENA")
    SPCLEAR
    surgicalPathology.Visible = True
    immunohistochemistry.Visible = False
    frm_hematopathology.Visible = False
    biochemistry.Visible = False
    bceltm.Visible = False
'    cardiology.Visible = False
    Microbiology.Visible = False
    CYTOLOGY.Visible = False
    fnac.Visible = False
    'cpl.Visible = False
    TUMOURMARKER.Visible = False
    hemcbc.Hide
    HEMBF.Hide
    HEMBMA.Hide
    'HLATYPING.Visible = False
    lbl_req.Left = 3100
    lbl_req.Caption = "REQUISITION - SURGICAL PATHOLOGY"
    pathology.Caption = "Pathology - Surgical Pathology"
    txt_dept.Text = "FZZ"
    
    ' PROCEDURE CODE FILLING IN THE COMBOBOX
'    If SELRINVESTFLAG = False Then
'        SELRINVEST_QR(0) = Trim(txt_modcode.Text)
'        Set SELRINVEST_RS = SELRINVEST_QR.OpenResultset(1, 3)
'        SELRINVESTFLAG = True
'    Else
'        SELRINVEST_QR(0) = Trim(txt_modcode.Text)
'        SELRINVEST_RS.MoreResults
'        SELRINVEST_RS.Requery
'    End If
'    If SELRINVEST_RS.RowCount > 0 Then
'        cmb_procedure.clear
'        cmb_procedurecode.clear
'        While Not SELRINVEST_RS.EOF
'            cmb_procedurecode.AddItem Trim(SELRINVEST_RS("ser_code"))
'            If SELMA101FLAG = False Then
'                SELMA101_QR(0) = Trim(SELRINVEST_RS("SER_CODE"))
'                Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
'                SELMA101FLAG = True
'             Else
'                SELMA101_QR(0) = SELRINVEST_RS("SER_CODE")
'                SELMA101_RS.MoreResults
'                SELMA101_RS.Requery
'            End If
'            If SELMA101_RS.RowCount > 0 Then
'                While Not SELMA101_RS.EOF
'                cmb_procedure.AddItem SELMA101_RS("SER_DET")
'                SELMA101_RS.MoveNext
'                Wend
'            End If
'            SELRINVEST_RS.MoveNext
'        Wend
'    End If

    ''ADDED BY MONO ON 02/05/2003
    
    If SELRINVEST1_FLAG = False Then
        SELRINVEST1_QR(0) = Trim(txt_modcode.Text)
        Set SELRINVEST1_RS = SELRINVEST1_QR.OpenResultset(1, 3)
        SELRINVEST1_FLAG = True
    Else
        SELRINVEST1_QR(0) = Trim(txt_modcode.Text)
        SELRINVEST1_RS.MoreResults
        SELRINVEST1_RS.Requery
    End If
    If SELRINVEST1_RS.RowCount > 0 Then
        cmb_procedure.clear
        cmb_procedurecode.clear
        While Not SELRINVEST1_RS.EOF
            cmb_procedurecode.AddItem Trim(SELRINVEST1_RS("ser_code"))
            'cmb_procedurecode.Text = Trim(SELRINVEST1_RS("ser_code"))
            If SELMA101_FLAG1 = False Then
                SELMA101_QR1(0) = Trim(SELRINVEST1_RS("SER_CODE"))
                Set SELMA101_RS1 = SELMA101_QR1.OpenResultset(1, 3)
                SELMA101_FLAG1 = True
             Else
                SELMA101_QR1(0) = SELRINVEST1_RS("SER_CODE")
                SELMA101_RS1.MoreResults
                SELMA101_RS1.Requery
            End If
            If SELMA101_RS1.RowCount > 0 Then
                While Not SELMA101_RS1.EOF
                cmb_procedure.AddItem SELMA101_RS1("SER_DET")
                'cmb_procedure.Text SELMA101_RS1("SER_DET")
                SELMA101_RS1.MoveNext
                Wend
            End If
            SELRINVEST1_RS.MoveNext
        Wend
    End If
    
    ' PROCEDURE CODE FILLING IN THE COMBOBOX
    ' SITES FILLING IN THE COMBOBOX
    If SELMA301FLAG = False Then
        Set SELMA301_RS = SELMA301_QR.OpenResultset(1, 3)
        SELMA301FLAG = True
    Else
        Set SELMA301_RS = SELMA301_QR.OpenResultset(1, 3)
        SELMA301_RS.MoreResults
        SELMA301_RS.Requery
    End If
    
    If SELMA301_RS.RowCount > 0 Then
        cmb_sites.clear
        While Not SELMA301_RS.EOF
            If Trim(SELMA301_RS("sitedesc")) <> "" Then
                cmb_sites.AddItem Trim(SELMA301_RS("SITEDESC")) & "$" & Trim(SELMA301_RS("SITECD"))
            End If
            SELMA301_RS.MoveNext
        Wend
    End If
    ' SITES FILLING IN THE COMBOBOX
    ' LATERALITY FILLING IN THE COMBOBOX
    If SELMA10122FLAG = False Then
        Set SELMA10122_RS = SELMA10122_QR.OpenResultset(1, 3)
        SELMA10122FLAG = True
    Else
    
        SELMA10122_RS.MoreResults
        SELMA10122_RS.Requery
    End If
    If SELMA10122_RS.RowCount > 0 Then
        cmb_laterality1.clear
        cmb_laterality.clear
        While Not SELMA10122_RS.EOF
        cmb_laterality1.AddItem Trim(SELMA10122_RS("ADJ_PRVDIA"))
        cmb_laterality.AddItem Trim(SELMA10122_RS("ADJ_PRVDIA"))
        SELMA10122_RS.MoveNext
        Wend
    End If
    ' LATERALITY FILLING IN THE COMBOBOX
  cmb_outblock.clear
  cmb_outslide.clear
  For I = 0 To 5
   cmb_outblock.AddItem I
   cmb_outslide.AddItem I
  Next I
  txt_spreqno.Enabled = False
  txt_sppathno.Enabled = False
  
    cmb_spfnac.clear
    cmb_spfnac.AddItem "YES"
    cmb_spfnac.AddItem "NO"
  ' FILLING THE PREVIOUSLY INVESTIGATED REQ NO.
   
'   If SELRREQUISITNFLAG = False Then
'        SELRREQUISITN_QR(0) = Msk_caseno.FormattedText
'        Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
'        SELRREQUISITNFLAG = True
'   Else
'
'
'        SELRREQUISITN_QR(0) = Msk_caseno.FormattedText
'        SELRREQUISITN_RS.MoreResults
'        SELRREQUISITN_RS.Requery
'
'    End If
'    If SELRREQUISITN_RS.RowCount > 0 Then
'        cmb_preinvest.clear
'        While Not SELRREQUISITN_RS.EOF
'            cmb_preinvest.AddItem SELRREQUISITN_RS("REQNNO")
'            SELRREQUISITN_RS.MoveNext
'        Wend
'        cmb_preinvest.ListIndex = 0
'    End If
''commented by mono
'            If seLdreqrsflag = False Then
'                    SELDREQQR(0) = Msk_caseno.FormattedText
'                    Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
'                    seLdreqrsflag = True
'                Else
'                    SELDREQQR(0) = Msk_caseno.FormattedText
'                    SELDREQRS.MoreResults
'                    SELDREQRS.Requery
'                End If
'                If SELDREQRS.RowCount > 0 Then
'                    cmb_preinvest.clear
'                    While Not SELDREQRS.EOF
'                        cmb_preinvest.AddItem SELDREQRS("PATHNO") & "$" & SELDREQRS("REQNO")
'                        SELDREQRS.MoveNext
'                    Wend
'                    cmb_preinvest.ListIndex = 0
'                End If

  ' FILLING THE PREVIOUSLY INVESTIGATED REQ NO.
  'FILLING THE OPERATIVE PROCEDURES
  If SELMA101FLAG1 = False Then
    Set SELMA101RS = selMA101QR.OpenResultset(1, 3)
    SELMA101FLAG1 = True
  Else
    SELMA101RS.MoreResults
    SELMA101RS.Requery
  End If
'  If SELMA101RS.RowCount > 0 Then
'    TXT_OPPROCEDURE.clear
'    While Not SELMA101RS.EOF
'        TXT_OPPROCEDURE.AddItem Trim(SELMA101RS("SER_DET"))
'        SELMA101RS.MoveNext
'    Wend
'  End If
  'FILLING THE OPERATIVE PROCEDURES
  'txt_sppathno.Visible = True
  'lbl_sppathno1.Visible = True
End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Public Sub FLDENABL(str1 As String)

If str1 = "INI" Then
    Msk_caseno.Enabled = False
    cmb_refdoc.Enabled = False
    cmb_reftodoc.Enabled = False
    txt_clinhist.Enabled = False
    txt_indctinvstg.Enabled = False
    cmb_dept.Enabled = False
    cmb_modality.Enabled = False
    txt_year.Enabled = False
    CMB_DURILLMONTH.Enabled = False
    If SELDISICU_FLAG = False Then
        SELDISICU_QR(0) = Trim(ipadd)
        Set SELDISICU_RS = SELDISICU_QR.OpenResultset(1, 3)
        SELDISICU_FLAG = True
    Else
        Set SELDISICU_RS = SELDISICU_QR.OpenResultset(1, 3)
        SELDISICU_QR(0) = Trim(ipadd)

        SELDISICU_RS.MoreResults
        SELDISICU_RS.Requery
    End If
    
'    If SELDISICU_QR(0) = "10.3.1.56" Then
'        lbl_rwd.Visible = True
'    Else
'        lbl_rwd.Visible = False
'    End If
    
    If SELDISICU_RS.RowCount > 0 Then
        If Trim(SELDISICU_RS(1)) = "10.3.1.56" Then
            lbl_rwd.Visible = True
            lbl_icureq.Visible = False
        Else
            lbl_rwd.Visible = False
            lbl_icureq.Visible = True
        End If
    End If
    
'    If Not SELDISICU_RS.EOF Then
'        lbl_icureq.Visible = True
'    Else
'        lbl_icureq.Visible = False
'    End If
    
End If
If str1 = "NEW" Then
    Msk_caseno.Enabled = True
    cmb_refdoc.Enabled = True
    cmb_reftodoc.Enabled = True
    txt_clinhist.Enabled = True
    txt_indctinvstg.Enabled = True
    cmb_dept.Enabled = True
    cmb_modality.Enabled = True
    txt_year.Enabled = True
    CMB_DURILLMONTH.Enabled = True
    If SELDISICU_FLAG = False Then
        SELDISICU_QR(0) = Trim(ipadd)
        Set SELDISICU_RS = SELDISICU_QR.OpenResultset(1, 3)
        SELDISICU_FLAG = True
    Else

        SELDISICU_QR(0) = Trim(ipadd)
        SELDISICU_RS.MoreResults
        SELDISICU_RS.Requery
    End If
    
    If SELDISICU_RS.RowCount > 0 Then
        If Trim(SELDISICU_RS(1)) = "10.3.1.56" Then
            lbl_rwd.Visible = True
            lbl_icureq.Visible = False
        Else
            lbl_rwd.Visible = False
            lbl_icureq.Visible = True
        End If
    End If
    
'    If SELDISICU_QR(0) = "10.3.1.56" Then
'        lbl_rwd.Visible = True
'    Else
'        lbl_rwd.Visible = False
'    End If
'
'    If Not SELDISICU_RS.EOF Then
'        lbl_icureq.Visible = True
'    Else
'        lbl_icureq.Visible = False
'    End If
    
End If
If str1 = "VIEW" Then
     Msk_caseno.Enabled = True
      cmb_refdoc.Enabled = False
      cmb_reftodoc.Enabled = False
      txt_clinhist.Enabled = False
      txt_indctinvstg.Enabled = False
      cmb_dept.Enabled = False
      cmb_modality.Enabled = False
      txt_year.Enabled = False
      CMB_DURILLMONTH.Enabled = False
      ICU = ""
    If SELDISICU_FLAG = False Then
        SELDISICU_QR(0) = Trim(ipadd)
        Set SELDISICU_RS = SELDISICU_QR.OpenResultset(1, 3)
        SELDISICU_FLAG = True
    Else
        SELDISICU_QR(0) = Trim(ipadd)
        SELDISICU_RS.MoreResults
        SELDISICU_RS.Requery
    End If
    
    If SELDISICU_RS.RowCount > 0 Then
        If Trim(SELDISICU_RS(1)) = "10.3.1.56" Then
            lbl_rwd.Visible = True
            lbl_icureq.Visible = False
        Else
            lbl_rwd.Visible = False
            lbl_icureq.Visible = True
        End If
    End If
    
'    If SELDISICU_QR(0) = "10.3.1.56" Then
'        lbl_rwd.Visible = True
'    Else
'        lbl_rwd.Visible = False
'    End If
'
'    If Not SELDISICU_RS.EOF Then
'        lbl_icureq.Visible = True
'    Else
'        lbl_icureq.Visible = False
'    End If
End If
End Sub


Public Sub REQNOAUTO()
On Error GoTo MSG:
If SELDAUTOREQNOFLAG = False Then
    SELDAUTOREQNO_QR(0) = Trim(txt_dept.Text)
    SELDAUTOREQNO_QR(1) = Trim(txt_modcode.Text)
    SELDAUTOREQNO_QR(2) = Right(DATE11, 2)
    Set SELDAUTOREQNO_RS = SELDAUTOREQNO_QR.OpenResultset(1, 3)
    SELDAUTOREQNOFLAG = True
 Else
    SELDAUTOREQNO_QR(0) = Trim(txt_dept.Text)
    SELDAUTOREQNO_QR(1) = Trim(txt_modcode.Text)
    SELDAUTOREQNO_QR(2) = Right(DATE11, 2)
    SELDAUTOREQNO_RS.MoreResults
    SELDAUTOREQNO_RS.Requery
End If
If SELDAUTOREQNO_RS.RowCount > 0 Then
    AUTOREQNO = SELDAUTOREQNO_RS("REQNSRNO") + 1
    UPDDAUTOREQNO_QR(0) = AUTOREQNO
    UPDDAUTOREQNO_QR(1) = Trim(txt_dept.Text)
    UPDDAUTOREQNO_QR(2) = Trim(txt_modcode.Text)
    UPDDAUTOREQNO_QR(3) = Trim(Right(DATE11, 2))
    UPDDAUTOREQNO_QR.Execute
Else
    AUTOREQNO = 1
    INSDAUTOREQNO_QR(0) = Trim(txt_dept.Text)
    INSDAUTOREQNO_QR(1) = Trim(txt_modcode.Text)
    INSDAUTOREQNO_QR(2) = Trim(Right(DATE11, 2))
    INSDAUTOREQNO_QR(3) = 1
    INSDAUTOREQNO_QR.Execute
End If
'***********padding the reqslno with leading zeros"****************
Dim reqslno1
reqslno1 = AUTOREQNO
While Len(CStr(reqslno1)) < 6
reqslno1 = "0" & reqslno1
Wend
AUTOREQNO = reqslno1
'******************************************************************
DAUTOREQNO = Trim(txt_dept.Text) & "/" & Trim(txt_modcode.Text) & "/" & Trim(Right(DATE11, 2)) & "/" & AUTOREQNO
If txt_modcode.Text = "SP" Then
    txt_spreqno = DAUTOREQNO
    If DPATHYEARFLAG = False Then
        DPATHYEAR_QR(0) = "A" & Right(DATE11, 4)
        Set DPATHYEAR_RS = DPATHYEAR_QR.OpenResultset(1, 3)
        DPATHYEARFLAG = True
    Else
        DPATHYEAR_QR(0) = "A" & Right(DATE11, 4)
        DPATHYEAR_RS.MoreResults
        DPATHYEAR_RS.Requery
    End If
    If DPATHYEAR_RS.RowCount > 0 Then
        PATH = DPATHYEAR_RS("LASTTWO")
    
    End If
    'MONO
    'txt_sppathno.Text = Right(DAUTOREQNO, 6) & "/" & Trim(PATH)
End If
'If txt_modcode.Text = "IH" Then  commented by candu 3/01/2002
'    txt_ihreqno.Text = DAUTOREQNO
'    If DPATHYEARFLAG = False Then
'        DPATHYEAR_QR(0) = "A" & Right(DATE11, 4)
'        Set DPATHYEAR_RS = DPATHYEAR_QR.OpenResultset(1, 3)
'        DPATHYEARFLAG = True
'    Else
'        DPATHYEAR_QR(0) = "A" & Right(DATE11, 4)
'        DPATHYEAR_RS.MoreResults
'        DPATHYEAR_RS.Requery
'    End If
'    If DPATHYEAR_RS.RowCount > 0 Then
'        PATH = DPATHYEAR_RS("LASTTWO")
'    End If
'    txt_ihpathno.Text = Right(DAUTOREQNO, 6) & "/" & Trim(PATH)
    
'End If
If txt_modcode.Text = "FA" Then
    txt_FNreqno.Text = DAUTOREQNO
End If
If Trim(txt_modcode.Text) = "CG" Or Trim(txt_modcode.Text) = "CN" Or Trim(txt_modcode.Text) = "PG" Or Trim(txt_modcode.Text) = "PN" Then
    txt_cyreqno1.Text = DAUTOREQNO
End If
If Trim(txt_modcode.Text) = "BC" Then
    txt_bcreqno.Text = DAUTOREQNO
End If
If Trim(txt_modcode.Text) = "EL" Then
    txt_bcreqno.Text = DAUTOREQNO
End If
If Trim(txt_modcode.Text) = "TM" Then
    txt_bcreqno.Text = DAUTOREQNO
End If
If Trim(txt_modcode.Text) = "BA" Or Trim(txt_modcode.Text) = "MB" Or Trim(txt_modcode.Text) = "SE" Or Trim(txt_modcode.Text) = "MC" Or Trim(txt_modcode.Text) = "CM" Or Trim(txt_modcode.Text) = "MO" Or Trim(txt_modcode.Text) = "XX" Then
    txt_mbreqno.Text = DAUTOREQNO
End If
If Trim(txt_modcode.Text) = "ED" Then
    txt_hpreqno.Text = DAUTOREQNO
End If
If Trim(txt_modcode.Text) = "2D" Or Trim(txt_modcode.Text) = "EK" Or Trim(txt_modcode.Text) = "PF" Or Trim(txt_modcode.Text) = "ST" Then
    txt_hpreqno.Text = DAUTOREQNO
End If
If Trim(txt_modcode.Text) = "HL" Or Trim(txt_modcode.Text) = "BM" Or Trim(txt_modcode.Text) = "BF" Then
    txt_bcreqno.Text = DAUTOREQNO

End If
If Trim(txt_modcode.Text) = "HT" Or Trim(txt_modcode.Text) = "SS" Or Trim(txt_modcode.Text) = "BP" Then
    txt_bcreqno.Text = DAUTOREQNO
End If
'  for cyotogenitics 13-04-2005
If Trim(txt_modcode.Text) = "CC" Or Trim(txt_modcode.Text) = "FI" Then
    txt_bcreqno.Text = DAUTOREQNO
End If
'  for cyotogenitics 13-04-2005
'for molecular pathology
If Trim(txt_modcode.Text) = "MP" Then
    txt_bcreqno.Text = DAUTOREQNO
End If


Exit Sub
MSG:
MsgBox Err.Description
End Sub

Public Sub PATHNOAUTO()
Dim PATH As String
On Error GoTo MSG:
If SELDAUTOPATHFLAG = False Then
    SELDAUTOPATH_QR(0) = Trim(txt_dept.Text)
    SELDAUTOPATH_QR(1) = Trim(txt_modcode.Text)
    SELDAUTOPATH_QR(2) = Trim(Right(DATE11, 2))
    Set SELDAUTOPATH_RS = SELDAUTOPATH_QR.OpenResultset(2, 3)
    SELDAUTOPATHFLAG = True
Else
    SELDAUTOPATH_QR(0) = Trim(txt_dept.Text)
    SELDAUTOPATH_QR(1) = Trim(txt_modcode.Text)
    SELDAUTOPATH_QR(2) = Trim(Right(DATE11, 2))
    SELDAUTOPATH_RS.MoreResults
    SELDAUTOPATH_RS.Requery
End If
If SELDAUTOPATH_RS.RowCount > 0 Then
    AUTOPATHNO = SELDAUTOPATH_RS("PATHNO") + 1
    UPDDAUTOPATH_QR(0) = AUTOPATHNO
    UPDDAUTOPATH_QR(1) = Trim(txt_dept.Text)
    UPDDAUTOPATH_QR(2) = Trim(txt_modcode.Text)
    UPDDAUTOPATH_QR(3) = Trim(Right(DATE11, 2))
    UPDDAUTOPATH_QR.Execute
Else
    AUTOPATHNO = 1
    INSDAUTOPATH_QR(0) = Trim(txt_dept.Text)
    INSDAUTOPATH_QR(1) = Trim(txt_modcode.Text)
    INSDAUTOPATH_QR(2) = Trim(Left(CASENO, 2))
    INSDAUTOPATH_QR(3) = Trim(Right(DATE11, 2))
    INSDAUTOPATH_QR(4) = 1
    INSDAUTOPATH_QR.Execute
End If
If DPATHYEARFLAG = False Then
    DPATHYEAR_QR(0) = "A" & Right(DATE11, 4)
    Set DPATHYEAR_RS = DPATHYEAR_QR.OpenResultset(1, 3)
    DPATHYEARFLAG = True
Else
    DPATHYEAR_QR(0) = "A" & Right(DATE11, 4)
    DPATHYEAR_RS.MoreResults
    DPATHYEAR_RS.Requery
End If
If DPATHYEAR_RS.RowCount > 0 Then
    PATH = DPATHYEAR_RS("LASTTWO")

End If
    '***********padding the reqslno with leading zeros"****************
Dim reqslno1
reqslno1 = AUTOPATHNO
While Len(CStr(reqslno1)) < 6
reqslno1 = "0" & reqslno1
Wend
DAUTOPATHNO = reqslno1
'******************************************************************
''MONO
'If txt_modcode.Text = "SP" Then
''txt_sppathno.Text = DAUTOPATHNO & "/" & "BS"
'txt_sppathno.Text = DAUTOPATHNO & "/" & PATH
'End If
If txt_modcode.Text = "SS" Then
    If BG1 = "Y" And Trim(TEXT1.Text) = "" Then
        TEXT1.Text = DAUTOPATHNO & "/" & Right(DATE11, 2)
    End If
End If
'If txt_modcode.Text = "SS" Then
'    Text1.Text = DAUTOPATHNO & "/" & Right(DATE11, 2)
'End If
'If txt_modcode.Text = "HT" Then
'    Text1.Text = DAUTOPATHNO & "/" & Right(DATE11, 2)
'End If


'If txt_modcode.Text = "IH" Then
''txt_ihpathno.Text = DAUTOPATHNO & "/" & "BS"
'txt_ihpathno.Text = DAUTOPATHNO & "/" & PATH
'End If
'If Trim(txt_modcode.Text) = "BA" Or Trim(txt_modcode.Text) = "MB" Or Trim(txt_modcode.Text) = "SE" Or Trim(txt_modcode.Text) = "MC" Or Trim(txt_modcode.Text) = "CM" Or Trim(txt_modcode.Text) = "MO" Or Trim(txt_modcode.Text) = "XX" Then
'    TXT_MBPATHNO.Text = DAUTOPATHNO & "/" & PATH
'End If
''dtm no. generation
'If txt_modcode.Text = "SS" Then
'   For i = 0 To CMB_TEXT.ListCount - 1
'        If CMB_TEXT.List(i) = "H001" Then
'            Text1.Text = DAUTOPATHNO & "/" & Right(DATE11, 2)
'        End If
'   Next i
'
'End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub
Public Sub NO_CHARGE()
'*****
On Error GoTo MSG:
No_Chargeflag = False
 
''If Left(CASENUM, 1) = "P" And Trim(lbl_deptcode.Caption) = "IZZ" Then
''  'pre_sercodech(0) = Trim(cbo_proccode.Text)
''  'Set selma101_RS = pre_sercodech.OpenResultset(1, 1)
''      'If selma101_RS.RowCount = 1 And selma101_RS("yn_flag") = "Y" Then
''          PRE_QUERY(0) = CASENUM
''          'PRE_QUERY(1) = selma101_RS("FROM_DATE")
''          'PRE_QUERY(2) = selma101_RS("TO_DATE")
''          Set PREV_RS = PRE_QUERY.OpenResultset(1, 1)
''            If PREV_RS.RowCount > 0 And Trim(lbl_deptcode.Caption) = "IZZ" Then
''                   No_Chargeflag = True
''              Else
''                   No_Chargeflag = False
''            End If
''      'Else
''
''      '    No_Chargeflag = False
''
''      'End If
''Else
'  SELMA101_QR(0) = Trim(cbo_proccode.Text)
'  SELMA101_RS.MoreResults
'  SELMA101_RS.Requery
'
'    If SELMA101_RS.RowCount > 0 Then
'        If Trim(lbl_catg.Caption) = "C" Then
'            If SELMA101_RS("C_CH") = 0 Then
'                No_Chargeflag = True
'            End If
'        ElseIf Trim(lbl_catg.Caption) = "NC" Then ' Or Opt_S.Value = True Then
'            If SELMA101_RS("NC_CH") = 0 Then
'                No_Chargeflag = True
'            End If
'        ElseIf Opt_S.Value = True Then
'                No_Chargeflag = True
'        End If
'    Else
'        MsgBox "There is NO  RECORD with this SERVICE CODE."
'    End If
''End If
'Exit Sub
'MSG:
'MsgBox Error(Err)



'*****
On Error GoTo MSG:
If Trim(txt_modcode.Text) = "CC" Or Trim(txt_modcode.Text) = "FI" Then
   cmb_procedurecode.Text = Trim(cmb_cyprocode.Text)
ElseIf Trim(txt_modcode.Text) = "MP" Then
   cmb_procedurecode.Text = Trim(molecularpathology.List_code.Text)
End If

No_Chargeflag = False
If SELMA101FLAG = False Then
    SELMA101_QR(0) = Trim(cmb_procedurecode.Text)
    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
    SELMA101FLAG = True
 Else
 '21-AUG-09 SHWETA
    SELMA101_QR(0) = Trim(cmb_procedurecode.Text)
    SELMA101_RS.MoreResults
    SELMA101_RS.Requery
End If

'SELMA101_QR(0) = Trim(cmb_procedurecode.Text)
'SELMA101_RS.MoreResults
'SELMA101_RS.Requery
If SELMA101_RS.RowCount > 0 Then
    If Trim(txt_catgory.Text) = "C" Then
        If SELMA101_RS("C_CH") = 0 Then
            No_Chargeflag = True
        End If
    ElseIf Trim(txt_catgory.Text) = "NC" Then
        If SELMA101_RS("NC_CH") = 0 Then
            No_Chargeflag = True
        End If
    End If
    If Trim(txt_catgory.Text) = "B" Then
        If SELMA101_RS("B_HOSP_CH") = 0 Then
            No_Chargeflag = True
        End If
    End If
    
Else
    'MsgBox "There is NO  RECORD with this SERVICE CODE." 'COMMENTED BY SHWETA 18-AUG-09
End If

Dim prev As String
prev = Mid(Msk_caseno.Text, 1, 1)
If No_Chargeflag = False And (prev = "P") Then
 'CHEKING EXTENSION OF I80 DAYS
    NOCHARG_QR(0) = Msk_caseno.FormattedText
    If NOCHARG_FLG = False Then
     Set NOCHARG_RS = NOCHARG_QR.OpenResultset(2, 3)
      NOCHARG_FLG = True
    Else
      NOCHARG_RS.MoreResults
      NOCHARG_RS.Requery
    End If
    If NOCHARG_RS.RowCount > 0 Then
    'CHECKING FREE CODE FOR PREVENTIVE ONCOLOGY
      PRFREE_QR(0) = Trim(cmb_procedurecode.Text)
        If PRFREE_FLG = False Then
           Set PRFREE_RS = PRFREE_QR.OpenResultset(2, 3)
           PRFREE_FLG = True
        Else
           PRFREE_RS.MoreResults
           PRFREE_RS.Requery
        End If
        If PRFREE_RS.RowCount > 0 Then
           No_Chargeflag = True
           NoMemo = True
        End If
    Else
       No_Chargeflag = False
    End If
End If

Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Public Sub SPASSIGN()
Dim durill As String
Dim I As Integer
On Error GoTo MSG:
'*** INSERTING INT0 RREQUISITION
INSRREQUISITN_QR(0) = req_chglen(Trim(txt_spreqno.Text))
INSRREQUISITN_QR(1) = GET_REVDATE(CStr(DATE11))
INSRREQUISITN_QR(2) = Trim(txt_time.Text)
INSRREQUISITN_QR(3) = Trim(Msk_caseno.FormattedText)
INSRREQUISITN_QR(4) = Trim(txt_dept.Text)
INSRREQUISITN_QR(5) = Trim(txt_modcode.Text)
INSRREQUISITN_QR(6) = Trim(txt_status.Text)
INSRREQUISITN_QR(7) = Trim(txt_catgory.Text)
INSRREQUISITN_QR(8) = Trim(cmb_procedurecode.Text)
INSRREQUISITN_QR(9) = UCase(Trim(cmb_refdoc.Text))
'INSRREQUISITN_QR(9) = Trim(txt_pridocser.Text)

'Added by Milind on 27-03-2006 for inserting date for the samples collected at outside
    If Check1.Value = 1 Then
        Dim Dt As String
        Dt = "18/02/1980"
        INSRREQUISITN_QR(15) = Format(CDate(Dt), "dd/mm/yyyy")
    Else
        INSRREQUISITN_QR(15) = Null
    End If
'upto here

    If SELRREQUISITNFLAG = False Then
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
        SELRREQUISITNFLAG = True
    Else
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        SELRREQUISITN_RS.MoreResults
        SELRREQUISITN_RS.Requery
    End If
    If SELRREQUISITN_RS.RowCount > 0 Then
        INSRREQUISITN_QR(10) = "YES"
    Else
        INSRREQUISITN_QR(10) = "NO"
    End If
If opt_patient.Value = True Then
    If Trim(cmb_procedurecode.Text) = "F101" Then
        NO_CHARGESP
    Else
        NO_CHARGE
    End If
    If No_Chargeflag = True Then
        INSRREQUISITN_QR(11) = "NOCHARGE"
        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
        INSRREQUISITN_QR(13) = "P"
        No_Chargeflag = False
    Else
        If Trim(cmb_vpp.Text) = "DER" Or Trim(cmb_vpp.Text) = "BAR" Or Trim(cmb_vpp.Text) = "DrOp" Or Trim(cmb_vpp.Text) = "VPP" Or Trim(cmb_vpp.Text) = "CHIP" Then
            
            If Trim(cmb_vpp.Text) = "DER" Then
                INSRREQUISITN_QR(11) = "DER"
                INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
                INSRREQUISITN_QR(13) = "P"
            End If
            If Trim(cmb_vpp.Text) = "BAR" Then
                INSRREQUISITN_QR(11) = "BAR"
                INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
                INSRREQUISITN_QR(13) = "P"
            End If
            If Trim(cmb_vpp.Text) = "DrOp" Then
                INSRREQUISITN_QR(11) = "DrOp"
                INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
                INSRREQUISITN_QR(13) = "P"
            End If
            If Trim(cmb_vpp.Text) = "VPP" Then
                INSRREQUISITN_QR(11) = "VPP"
                INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
                INSRREQUISITN_QR(13) = Null    ' MADE  NULL FROM P
            End If
''''            If Trim(cmb_vpp.Text) = "CHIP" Then
''''                INSRREQUISITN_QR(11) = "CHIP"
''''                INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
''''                INSRREQUISITN_QR(13) = "P"    ' MADE  NULL FROM P
''''            End If
            
            
                    '''ICU
'        ElseIf Trim(CMB_ICU1.Text) <> "" Then
'            If Trim(CMB_ICU1.Text) = "YES" Then
'                INSRREQUISITN_QR(11) = Null
'                INSRREQUISITN_QR(12) = Null
'                INSRREQUISITN_QR(13) = "P"
'            End If
        Else
            INSRREQUISITN_QR(11) = Null
            INSRREQUISITN_QR(12) = Null
            INSRREQUISITN_QR(13) = Null
        End If
    End If
End If
If opt_staff.Value = True Then
        INSRREQUISITN_QR(11) = "NOCHARGE"
        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
        INSRREQUISITN_QR(13) = "P"
End If
If Trim(cmb_vpp.Text) = "NONE" Or Len(Trim(cmb_vpp.Text)) = 0 Then
    INSRREQUISITN_QR(14) = 0
End If
If Trim(cmb_vpp.Text) = "VPP" Then
    INSRREQUISITN_QR(14) = 1
End If
If Trim(cmb_vpp.Text) = "DER" Then
    INSRREQUISITN_QR(14) = 2
End If
If Trim(cmb_vpp.Text) = "BAR" Then
    INSRREQUISITN_QR(14) = 3
End If
If Trim(cmb_vpp.Text) = "DrOp" Then
    INSRREQUISITN_QR(14) = 4
End If
'''If Trim(cmb_vpp.Text) = "CHIP" Then
'''    INSRREQUISITN_QR(14) = 5
'''End If
INSRREQUISITN_QR.Execute
    '*** INSERTING INT0 RREQUISITION
    '**** INSERTING INTO DREQUISITION

INSDREQUISITN_QR(0) = req_chglen(Trim(txt_spreqno.Text))
INSDREQUISITN_QR(1) = " "
'MONO
'INSDREQUISITN_QR(1) = Trim(txt_sppathno.Text)
INSDREQUISITN_QR(2) = Trim(Msk_caseno.FormattedText)
                            
 I = InStr(cmb_sites, "$")
 If I = 0 Then
                            
               If sitedupflag = False Then
                    sitedupqr(0) = UCase(Trim(cmb_sites.Text))
                    Set siteduprs = sitedupqr.OpenResultset(1, 3)
                    sitedupflag = True
               Else
                    sitedupqr(0) = UCase(Trim(cmb_sites.Text))
                    siteduprs.MoreResults
                    siteduprs.Requery
                                   
               End If
                            
        If siteduprs.RowCount > 0 Then
                INSDREQUISITN_QR(3) = siteduprs("SITECD")
        Else
                If sitedup1flag = False Then
                    'sitedupqr(0) = Trim(cmb_sites.Text)
                    Set siteduprs1 = sitedupqr1.OpenResultset(1, 3)
                    sitedup1flag = True
                Else
                    siteduprs1.MoreResults
                    siteduprs1.Requery
                End If
                If (siteduprs1.RowCount > 0) Then
                    Dim a As Integer
                    If Len(Trim(cmb_sites.Text)) > 0 Then
                                     
                    a = siteduprs1(0)
                    a = a + 1
                    siteinsertqr(0) = a
                    siteinsertqr(1) = UCase(Trim(cmb_sites.Text))
                    siteinsertqr.Execute
                    INSDREQUISITN_QR(3) = a
                    Else
                        INSDREQUISITN_QR(3) = "9999"
                    End If
                 Else
                    
                    a = 5000
                    siteinsertqr(0) = a
                    siteinsertqr(1) = UCase(Trim(cmb_sites.Text))
                    siteinsertqr.Execute
                    INSDREQUISITN_QR(3) = a
                 End If
         End If
  Else
        INSDREQUISITN_QR(3) = Mid(cmb_sites, I + 1, 4)
  End If
INSDREQUISITN_QR(4) = Trim(cmb_laterality1.Text)
dur_ill = txt_year & "/" & Trim(CMB_DURILLMONTH.Text)
INSDREQUISITN_QR(5) = dur_ill
INSDREQUISITN_QR(6) = Trim(txt_metastasis.Text)
INSDREQUISITN_QR(7) = Trim(cmb_laterality.Text)
INSDREQUISITN_QR(8) = Trim(txt_outspecimen.Text)
INSDREQUISITN_QR(9) = Trim(cmb_outblock.Text)
INSDREQUISITN_QR(10) = Trim(cmb_outslide.Text)
INSDREQUISITN_QR(11) = Trim(cmb_frozensec.Text)
If cmb_frozensec.Text = "YES" Then
    INSDREQUISITN_QR(12) = GET_REVDATE(Trim(MSK_FROSECDT.FormattedText))
Else
    INSDREQUISITN_QR(12) = Null
End If
INSDREQUISITN_QR(13) = Null
INSDREQUISITN_QR(14) = Trim(cmb_refdoc.Text)
INSDREQUISITN_QR(15) = Trim(cmb_reftodoc.Text)
INSDREQUISITN_QR(16) = Trim(TXT_OPPROCEDURE.Text)
If Trim(txt_status.Text) = "IP" Then
    INSDREQUISITN_QR(17) = Trim(txt_ward.Text)
    INSDREQUISITN_QR(18) = Trim(txt_bedno.Text)
Else
    INSDREQUISITN_QR(17) = Null
    INSDREQUISITN_QR(18) = Null
End If
INSDREQUISITN_QR(19) = " "
If Len(Trim(txt_nmeterial.Text)) > 0 Then
    INSDREQUISITN_QR(20) = Null
    INSDREQUISITN_QR(20) = txt_nmeterial
Else
    INSDREQUISITN_QR(20) = Null
    INSDREQUISITN_QR(20) = " "
End If
If Len(Trim(cmb_spfnac.Text)) = 0 Then
    INSDREQUISITN_QR(21) = Null
Else
    INSDREQUISITN_QR(21) = cmb_spfnac.Text
End If
INSDREQUISITN_QR.Execute
'**** INSERTING INTO DREQUISITION

 '*** UPDATING THE PROVISIONAL DIAGNOSIS
 
If opt_patient.Value = True Then
If SELTA1012FLAG = False Then
    selta1012_qr(0) = Msk_caseno.FormattedText
    Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
    SELTA1012FLAG = True
Else
     selta1012_qr(0) = Msk_caseno.FormattedText
     selta1012_rs.MoreResults
     selta1012_rs.Requery
End If
If selta1012_rs.RowCount > 0 Then
    UPDTA1012_QR(0) = Trim(txt_indctinvstg.Text)
    UPDTA1012_QR(1) = Msk_caseno.FormattedText
    UPDTA1012_QR.Execute
Else
    insta1012_qr(0) = Msk_caseno.FormattedText
    insta1012_qr(1) = Trim(txt_indctinvstg.Text)
    insta1012_qr.Execute
End If
End If
'*** UPDATING THE PROVISIONAL DIAGNOSIS
'*** INSERTING INTO DMULPROCE FOR COLLECTION AND DEPT COLLECTION SCREEN
''    INSDMULPROCE_QR(0) = req_chglen(txt_spreqno.Text)
''    INSDMULPROCE_QR(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
''    INSDMULPROCE_QR(2) = Msk_caseno.FormattedText
''    INSDMULPROCE_QR(3) = cmb_procedurecode.Text
''    INSDMULPROCE_QR.Execute

'***INSERTING INTO DMULPROCE FOR COLLECTION AND DEPT COLLECTION SCREEN
'***INSERTING INTO DREQSTAFFD TABLE
If opt_staff.Value = True Then
    INSDREQSTAFFD_QR(0) = req_chglen(txt_spreqno.Text)
    INSDREQSTAFFD_QR(1) = Msk_caseno.Text
    INSDREQSTAFFD_QR(2) = Trim(txt_name.Text)
    INSDREQSTAFFD_QR(3) = Trim(txt_Age)
    INSDREQSTAFFD_QR(4) = Trim(txt_sex)
    If Len(Trim(cmb_dependents.Text)) = 0 Then
        INSDREQSTAFFD_QR(5) = "SELF"
    Else
        I = InStr(cmb_dependents, "$")
        INSDREQSTAFFD_QR(5) = Mid(cmb_dependents, I + 1, 10)
    End If
    INSDREQSTAFFD_QR.Execute
End If
'***INSERTING INTO DREQSTAFFD TABLE
    If SELCLINHISTFLAG = False Then
        SELCLINHIST_QR(0) = Msk_caseno.FormattedText
        Set SELCLINHIST_RS = SELCLINHIST_QR.OpenResultset(1, 3)
        SELCLINHISTFLAG = True
    Else
        SELCLINHIST_QR(0) = Msk_caseno.FormattedText
        SELCLINHIST_RS.MoreResults
        SELCLINHIST_RS.Requery
    End If
'    If SELCLINHIST_RS.RowCount > 0 Then
'        updclinhist_qr(0) = Trim(txt_clinhist.Text)
'        updclinhist_qr(1) = Trim(Msk_caseno.Text)
'        updclinhist_qr.Execute
'    Else
        insclinhist_qr(0) = Trim(Msk_caseno.FormattedText)
        insclinhist_qr(1) = GET_REVDATE(Trim(DATE11))
        insclinhist_qr(2) = Trim(time1)
        If Trim(txt_clinhist.Text) = "NYD" Then
            insclinhist_qr(3) = "NYD"
        Else
            insclinhist_qr(3) = Trim(txt_clinhist.Text)
        End If
        insclinhist_qr.Execute
    'End If

Data1.UpdateRecord
Data1.Recordset.Bookmark = Data1.Recordset.LastModified
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Public Sub VIEWRETRIVE()
Dim I As Integer
On Error GoTo MSG:
If viewflag = True Then
    bfmode = "VIEW"
    
      If opt_patient.Value = True Then
           txt_name.Text = Trim(SELTA101_RS("TAG")) & " " & Trim(SELTA101_RS("FIRST_NAME")) & " " & Trim(SELTA101_RS("MID_NAME")) & " " & Trim(SELTA101_RS("LAST_NAME"))
           txt_sex = Trim(SELTA101_RS("SEX"))
               If Len(Trim(SELTA101_RS("BIRTH_DATE"))) <> 0 Then
                   YEAR1 = DateDiff("YYYY", SELTA101_RS("BIRTH_DATE"), DATE11)
                   PATYPE = Trim(SELTA101_RS("PAT_TYPE"))
                   txt_Age = YEAR1
               End If
           txt_catgory = Trim(SELTA101_RS("PAT_CATG"))
'           If selma104flag = False Then
'               selma104_qr(0) = msk_caseno.FormattedText
'               Set selma104_rs = selma104_qr.OpenResultset(1, 3)
'               selma104flag = True
'           Else
'               selma104_qr(0) = msk_caseno.FormattedText
'               selma104_rs.MoreResults
'               selma104_rs.Requery
'           End If
            If SELICU_FLAG = False Then
                SELICU_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                Set SELICU_RS = SELICU_QR.OpenResultset(1, 3)
                SELICU_FLAG = True
            Else
                SELICU_QR(0) = req_chglen(Trim(cmb_reqno.Text))
                SELICU_RS.MoreResults
                SELICU_RS.Requery
            End If
            
            If SELICU_RS.RowCount > 0 Then
                If Trim(SELICU_RS("ICDS2")) = "YES" Then
                    lbl_icureq.Visible = True
                    MCICU = "ICU"
                ElseIf Trim(SELICU_RS("ICDS2")) = "YR" Then
                    lbl_rwd.Visible = True
                    'MCRWD = "RECOVERY WARD"
                End If
            Else
                lbl_icureq.Visible = False
                MCICU = ""
                lbl_rwd.Visible = False
                MCRWD = ""
            End If
            
            If SELRREQUISITNFLAG2 = False Then
                SELRREQUISITN_QR2(0) = req_chglen(Trim(cmb_reqno.Text))
                Set SELRREQUISITN_RS2 = SELRREQUISITN_QR2.OpenResultset(1, 3)
                SELRREQUISITNFLAG2 = True
           Else
                SELRREQUISITN_QR2(0) = req_chglen(Trim(cmb_reqno.Text))
                SELRREQUISITN_RS2.MoreResults
                SELRREQUISITN_RS2.Requery
           End If
           If SELDREQUISITNFLAG = False Then
                seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                SELDREQUISITNFLAG = False
           Else
                seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                seldrequisitn_rs.MoreResults
                seldrequisitn_rs.Requery
           End If
           
           If SELRREQUISITN_RS2.RowCount > 0 Then
               If SELRREQUISITN_RS2("PATNSTAT") = "IP" Then
                txt_status = "IP"
                txt_bedno.Visible = True
                txt_ward.Visible = True
                lbl_bedno.Visible = True
                lbl_ward.Visible = True
                txt_bedno.Text = seldrequisitn_rs("BEDNO")
                txt_ward.Text = seldrequisitn_rs("WARD_NAME")
            Else
                txt_status = "OP"
                txt_bedno.Visible = False
                txt_ward.Visible = False
                lbl_bedno.Visible = False
                lbl_ward.Visible = False
            End If
           End If
           txt_pridocser.Text = Trim(SELTA101_RS("UD_NAME"))
           If SELTA1011FLAG = False Then
               SELTA1011_QR(0) = Trim(Msk_caseno.FormattedText)
               Set SELTA1011_RS = SELTA1011_QR.OpenResultset(2, 3)
               SELTA1011FLAG = True
           Else
               SELTA1011_QR(0) = Trim(Msk_caseno.FormattedText)
               SELTA1011_RS.MoreResults
               SELTA1011_RS.Requery
           End If
           If SELTA1011_RS.RowCount > 0 Then
               TXT_REF_DOC = Trim(SELTA1011_RS("REF_BY_DOC"))
           Else
               TXT_REF_DOC = ""
           End If
           If SELCLINHISTFLAG = False Then
               SELCLINHIST_QR(0) = Trim(Msk_caseno.FormattedText)
               Set SELCLINHIST_RS = SELCLINHIST_QR.OpenResultset(1, 3)
               SELCLINHISTFLAG = True
           Else
               SELCLINHIST_QR(0) = Trim(Msk_caseno.FormattedText)
               SELCLINHIST_RS.MoreResults
               SELCLINHIST_RS.Requery
           End If
           If SELCLINHIST_RS.RowCount > 0 Then
               SELCLINHIST_RS.MoveLast
               txt_clinhist.Text = Trim(SELCLINHIST_RS("CLINHIST"))
           Else
               txt_clinhist.Text = "NYD"
           End If
           If SELTA1012FLAG = False Then
                selta1012_qr(0) = Msk_caseno.FormattedText
                Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
                SELTA1012FLAG = False
           Else
                selta1012_qr(0) = Msk_caseno.FormattedText
                selta1012_rs.MoreResults
                selta1012_rs.Requery
                SELTA1012FLAG = True
           End If
           If selta1012_rs.RowCount > 0 Then
                txt_indctinvstg.Text = selta1012_rs("PRVDIAGDES")
           End If
           If SELRREQUISITNFLAG2 = False Then
                SELRREQUISITN_QR2(0) = req_chglen(Trim(cmb_reqno.Text))
                Set SELRREQUISITN_RS2 = SELRREQUISITN_QR2.OpenResultset(1, 3)
                SELRREQUISITNFLAG2 = True
           Else
                SELRREQUISITN_QR2(0) = req_chglen(Trim(cmb_reqno.Text))
                SELRREQUISITN_RS2.MoreResults
                SELRREQUISITN_RS2.Requery
           End If
           If SELRREQUISITN_RS2.RowCount > 0 Then
                txt_modcode.Text = SELRREQUISITN_RS2("MODALITYCD")
                txt_dept.Text = SELRREQUISITN_RS2("DEPT_CODE")
                
           End If
           
           If SELRMODALTIYFLAG2 = False Then
                SELRMODALITY_QR2(0) = Trim(txt_modcode.Text)
                Set SELRMODALITY_RS2 = SELRMODALITY_QR2.OpenResultset(1, 3)
                SELRMODALTIYFLAG2 = True
            Else
                SELRMODALITY_QR2(0) = Trim(txt_modcode.Text)
                SELRMODALITY_RS2.MoreResults
                SELRMODALITY_RS2.Requery
           End If
           If SELRMODALITY_RS2.RowCount > 0 Then
                cmb_modality.Text = Trim(SELRMODALITY_RS2("MODLITYDES"))
           End If
           If SELMA106FLAG2 = False Then
                SELMA106_QR2(0) = Trim(txt_dept.Text)
                Set SELMA106_RS2 = SELMA106_QR2.OpenResultset(1, 3)
                SELMA106FLAG2 = True
           Else
                SELMA106_QR2(0) = Trim(txt_dept.Text)
                SELMA106_RS2.MoreResults
                SELMA106_RS2.Requery
           End If
           If SELMA106_RS2.RowCount > 0 Then
''                If txt_modcode = "IH" Then
''                    cmb_dept.Text = "PATHOLOGY"
''                Else
                    cmb_dept.Text = Trim(SELMA106_RS2("DEPT_NAME"))
''                End If
           End If
           If SELDREQUISITNFLAG = False Then
                seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                SELDREQUISITNFLAG = False
           Else
                seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                seldrequisitn_rs.MoreResults
                seldrequisitn_rs.Requery
           End If
           If seldrequisitn_rs.RowCount > 0 Then
                
                'txt_pathno.Text = seldrequisitn_rs("PATHNO")
            End If
          '****** SP
          If Trim(txt_modcode.Text) = "SP" Then
                SPVIEW
          End If
''          If txt_modcode.Text = "IH" Then
''              IHVIEW
''          End If
    '**** FA FILLING
          If Trim(txt_modcode.Text) = "FA" Then
                FNACVIEW
          End If
   '***** CG OR CN
           If Trim(txt_modcode.Text) = "CG" Or Trim(txt_modcode.Text) = "CN" Or Trim(txt_modcode.Text) = "PG" Or Trim(txt_modcode.Text) = "PN" Then
                CGVIEW
           End If
   '***** CG OR CN
   '***** BIOCHEMISTRY
           If Trim(txt_modcode.Text) = "BC" Then
                BIOVIEW
           End If
           If Trim(txt_modcode.Text) = "EL" Then
                EMERVIEW
           End If
           If Trim(txt_modcode.Text) = "TM" Then
                TUMOURVIEW
           End If
   '***** BIOCHEMISTRY
   '***** MICROBIOLOGY
        If Trim(txt_modcode.Text) = "BA" Or Trim(txt_modcode.Text) = "MB" Or Trim(txt_modcode.Text) = "SE" Or Trim(txt_modcode.Text) = "MC" Or Trim(txt_modcode.Text) = "CM" Or Trim(txt_modcode.Text) = "MO" Or Trim(txt_modcode.Text) = "XX" Then
                MBVIEW
        End If
   '***** MICROBIOLOGY
   '***** ENDOSCOPY
            If Trim(txt_modcode.Text) = "ED" Then
                ENDOSVIEW
            End If
   
   '***** ENDOSCOPY
   '***** CARDIOLOGY
           If Trim(txt_modcode.Text) = "2D" Or Trim(txt_modcode.Text) = "EK" Or Trim(txt_modcode.Text) = "PF" Or Trim(txt_modcode.Text) = "ST" Then
                CARDVIEW
           End If
   '***** CARDIOLOGY
   '***** HAEMATOLOGY
           If Trim(txt_modcode.Text) = "HL" Then
                HEMVIEW
           End If
   '***** HAEMATOLOGY
   '***** Bone marrow
        If Trim(txt_modcode.Text) = "BM" Then
                BONEMVIEW
        End If
       
   '***** BONE MARROW
   '**** BODY FLUIDS
          If Trim(txt_modcode.Text) = "BF" Then
                BODYFVIEW
          End If
   
   '**** BODY FLUIDS
   '***** HLA TYPING
           If Trim(txt_modcode.Text) = "HT" Then
                HLAVIEW
           End If
   '***** HLA TYPING
   '***** DTM LAB
           If Trim(txt_modcode.Text) = "SS" Then
                dtmview
           End If
   '***** DTM LAB
   '***** BLOOD PRODUCTS
          If Trim(txt_modcode.Text) = "BP" Then
               BPVIEW
          End If
   '***** BLOOD PRODUCTS
   '***** CYTOGENITICS
          If Trim(txt_modcode.Text) = "CC" Or Trim(txt_modcode.Text) = "FI" Then
               cyview
               cmb_cyprocedure.Text = seldrequisitn_rs("OPPROCEDUR")
               
          End If
   '***** CYTOGENITICS
    End If
    If opt_staff.Value = True Then
             If SELSTAFFINFOFLAG = False Then
              SELSTAFFINFO_QR(0) = req_chglen(cmb_reqno)
              Set SELSTAFFINFO_RS = SELSTAFFINFO_QR.OpenResultset(2, 3)
              SELSTAFFINFOFLAG = True
             Else
              SELSTAFFINFO_RS.MoreResults
              SELSTAFFINFO_RS.Requery
             End If
             If SELSTAFFINFO_RS.RowCount > 0 Then
                 txt_name.Text = Trim(SELSTAFFINFO_RS("pname")) '& " " & Trim(SELSTAFFINFO_RS("FNAME")) & " " & Trim(SELSTAFFINFO_RS("MNAME")) & " " & Trim(SELSTAFFINFO_RS("LNAME"))
                 txt_sex = Trim(SELSTAFFINFO_RS("sex"))
                 txt_Age = Trim(SELSTAFFINFO_RS("age")) 'DateDiff("YYYY", SELSTAFFINFO_RS("DOB"), DATE11)
                 txt_status = "ST"
            End If
           
           If selma104flag = False Then
               selma104_qr(0) = Msk_caseno.FormattedText
               Set selma104_rs = selma104_qr.OpenResultset(1, 3)
               selma104flag = True
           Else
               selma104_qr(0) = Msk_caseno.FormattedText
               selma104_rs.MoreResults
               selma104_rs.Requery
           End If
           If selma104_rs.RowCount > 0 Then
               txt_status = "IP"
               txt_bedno.Text = Trim(selma104_rs("BED_NO"))
               txt_ward.Text = Trim(selma104_rs("WARD_NAME"))
           Else
               txt_status = "ST"
               txt_bedno.Text = ""
               txt_ward.Text = ""
           End If
           'txt_pridocser.Text = Trim(selta101_rs("UD_NAME"))
'           If SELTA1011FLAG = False Then
'               SELTA1011_QR(0) = Msk_caseno.FormattedText
'               Set SELTA1011_RS = SELTA1011_QR.OpenResultset(2, 3)
'               SELTA1011FLAG = True
'           Else
'               SELTA1011_QR(0) = Msk_caseno.FormattedText
'               SELTA1011_RS.MoreResults
'               SELTA1011_RS.Requery
'           End If
'           If SELTA1011_RS.RowCount > 0 Then
'               TXT_REF_DOC = Trim(SELTA1011_RS("REF_BY_DOC"))
'           Else
'               TXT_REF_DOC = ""
'           End If
           If SELCLINHISTFLAG = False Then
               SELCLINHIST_QR(0) = Trim(Msk_caseno.FormattedText)
               Set SELCLINHIST_RS = SELCLINHIST_QR.OpenResultset(1, 3)
               SELCLINHISTFLAG = True
           Else
               SELCLINHIST_QR(0) = Trim(Msk_caseno.FormattedText)
               SELCLINHIST_RS.MoreResults
               SELCLINHIST_RS.Requery
           End If
           If SELCLINHIST_RS.RowCount > 0 Then
               'SELCLINHIST_RS.MoveNext
               txt_clinhist.Text = Trim(SELCLINHIST_RS("CLINHIST"))
           Else
               txt_clinhist.Text = "NYD"
           End If
'           If SELTA1012FLAG = False Then
'                selta1012_qr(0) = Msk_caseno.FormattedText
'                Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
'                SELTA1012FLAG = False
'           Else
'                selta1012_qr(0) = Msk_caseno.FormattedText
'                selta1012_rs.MoreResults
'                SELTA1012FLAG = True
'           End If
'           If selta1012_rs.RowCount > 0 Then
'                txt_indctinvstg.Text = selta1012_rs("PRVDIAGDES")
'           End If
           If SELRREQUISITNFLAG2 = False Then
                SELRREQUISITN_QR2(0) = req_chglen(Trim(cmb_reqno.Text))
                Set SELRREQUISITN_RS2 = SELRREQUISITN_QR2.OpenResultset(1, 3)
                SELRREQUISITNFLAG2 = True
           Else
                SELRREQUISITN_QR2(0) = req_chglen(Trim(cmb_reqno.Text))
                SELRREQUISITN_RS2.MoreResults
                SELRREQUISITN_RS2.Requery
           End If
           If SELRREQUISITN_RS2.RowCount > 0 Then
                txt_modcode.Text = SELRREQUISITN_RS2("MODALITYCD")
                txt_dept.Text = SELRREQUISITN_RS2("DEPT_CODE")
           End If
           
           If SELRMODALTIYFLAG2 = False Then
                SELRMODALITY_QR2(0) = Trim(txt_modcode.Text)
                Set SELRMODALITY_RS2 = SELRMODALITY_QR2.OpenResultset(1, 3)
                SELRMODALTIYFLAG2 = True
            Else
                SELRMODALITY_QR2(0) = Trim(txt_modcode.Text)
                SELRMODALITY_RS2.MoreResults
                SELRMODALITY_RS2.Requery
           End If
           If SELRMODALITY_RS2.RowCount > 0 Then
                cmb_modality.Text = Trim(SELRMODALITY_RS2("MODLITYDES"))
           End If
           If SELMA106FLAG2 = False Then
                SELMA106_QR2(0) = Trim(txt_dept.Text)
                Set SELMA106_RS2 = SELMA106_QR2.OpenResultset(1, 3)
                SELMA106FLAG2 = True
           Else
                SELMA106_QR2(0) = Trim(txt_dept.Text)
                SELMA106_RS2.MoreResults
                SELMA106_RS2.Requery
           End If
           If SELMA106_RS2.RowCount > 0 Then
''                If txt_modcode = "IH" Then
''                    cmb_dept.Text = "PATHOLOGY"
''                Else
                    cmb_dept.Text = Trim(SELMA106_RS2("DEPT_NAME"))
''                End If
           End If
           If SELDREQUISITNFLAG = False Then
                seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                SELDREQUISITNFLAG = False
           Else
                seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                seldrequisitn_rs.MoreResults
                seldrequisitn_rs.Requery
           End If
           If seldrequisitn_rs.RowCount > 0 Then
                txt_pathno.Text = seldrequisitn_rs("PATHNO")
            End If
         '*****
          
               If Trim(txt_modcode.Text) = "SP" Then
          
                surgicalPathology.Visible = True
                cmb_procedurecode.Text = SELRREQUISITN_RS2("SER_CODE")
                If SELRREQUISITN_RS2.RowCount > 0 Then
                    If SELRREQUISITN_RS2("RECORDNO") = 1 Then
                    cmb_vpp.Text = "VPP"
                    End If
                    If SELRREQUISITN_RS2("RECORDNO") = 2 Then
                    cmb_vpp.Text = "DER"
                    End If
                    If SELRREQUISITN_RS2("RECORDNO") = 3 Then
                    cmb_vpp.Text = "BAR"
                    End If
                    If SELRREQUISITN_RS2("RECORDNO") = 4 Then
                    cmb_vpp.Text = "DrOp"
                    End If
''                    If SELRREQUISITN_RS2("RECORDNO") = 5 Then
''                    cmb_vpp.Text = "CHIP"
''                    End If
                    
                End If
                
                If SELMA101FLAG = False Then
                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS2("SER_CODE"))
                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                    SELMA101FLAG = True
                Else
                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS2("SER_CODE"))
                    SELMA101_RS.MoreResults
                    SELMA101_RS.Requery
                End If
                If SELMA101_RS.RowCount > 0 Then
                    cmb_procedure.Text = SELMA101_RS("SER_DET")
                End If
                If SELDREQUISITNFLAG = False Then
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                    SELDREQUISITNFLAG = True
                Else
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    seldrequisitn_rs.MoreResults
                    seldrequisitn_rs.Requery
                End If
                If seldrequisitn_rs.RowCount > 0 Then
                    If SELMA301FLAG = False Then
                        SELMA301QR(0) = seldrequisitn_rs("SITE")
                        Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
                        SELMA301FLAG = True
                    Else
                        SELMA301QR(0) = seldrequisitn_rs("SITE")
                        Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
                        SELMA301RS.MoreResults
                        SELMA301RS.Requery
                
                    End If
                    If SELMA301RS.RowCount > 0 Then
                        If Trim(SELMA301RS("SITEDESC")) = "site not specified/recorded" Then
                            cmb_sites.Text = ""
                        Else
                            cmb_sites.Text = SELMA301RS("SITEDESC")
                        End If
                        Else
'                        venkat*********
'                            Dim viewsitedupqr As rdoQuery
'                            Dim viewsiteduprs As rdoResultset
'                            Dim viewstr As String
'                            viewstr = "select sitecd,sitedesc from dislib.sitedup where sitecd = ?"
'                            Set viewsitedupqr = rdb.CreateQuery(viewsitedupqr1, viewstr)
'                            viewsitedupqr(0) = seldrequisitn_rs("SITE")
'                            Set viewsiteduprs = viewsitedupqr.OpenResultset(1, 3)
'                            cmb_sites.Text = viewsiteduprs("SITEDESC")
                    End If
                    cmb_laterality1.Text = seldrequisitn_rs("LATERALITS")
                    txt_metastasis.Text = seldrequisitn_rs("METASTASIS")
                    txt_outspecimen.Text = seldrequisitn_rs("OUTSSPECIM")
                    cmb_outblock.Text = seldrequisitn_rs("OUTBLOCK")
                    cmb_outslide.Text = seldrequisitn_rs("OUTSLIDES")
                    cmb_frozensec.Text = seldrequisitn_rs("FROZENSEC")
                    If IsNull(Trim(seldrequisitn_rs("FROZENDATE"))) Then
                    MSK_FROSECDT.Text = "__/__/____"
                    Else
                    MSK_FROSECDT.Text = seldrequisitn_rs("FROZENDATE")
                    End If
                    cmb_laterality = seldrequisitn_rs("LATERALITM")
                    TXT_OPPROCEDURE.Text = seldrequisitn_rs("OPPROCEDUR")
                    cmb_preinvest.Text = Trim(SELRREQUISITN_RS2("PREVEXAM"))
                    cmb_refdoc.Text = Trim(seldrequisitn_rs("ReFDOC"))
                    cmb_reftodoc.Text = Trim(seldrequisitn_rs("REFTODOC"))
                    If seldrequisitn_rs.EOF Then
                        cmb_spfnac.Text = " " & Trim(seldrequisitn_rs("LABNO"))
                    Else
                        cmb_spfnac.Text = ""
                    End If
                    If seldrequisitn_rs.EOF = False Then
                        txt_nmeterial.Text = " " & Trim(seldrequisitn_rs("NMETERIALS"))
                    Else
                        txt_nmeterial.Text = ""
                    End If
                    If Len(Trim(seldrequisitn_rs("DURILLNESS"))) > 1 Then
                    
                    I = InStr(seldrequisitn_rs("DURILLNESS"), "/")
                    txt_year.Text = Mid(seldrequisitn_rs("DURILLNESS"), 1, I - 1)
                    CMB_DURILLMONTH.Text = Mid(seldrequisitn_rs("DURILLNESS"), I + 1, 2)
                    
                    End If
                    SPENABLE ("DIS")
                    lbl_spreqno1.Visible = False
                    lbl_sppathno.Visible = True
                    txt_sppathno.Visible = False
                    txt_spreqno.Visible = False
                    ' Previously Investigated requisition nos
'                    If SELRREQFLAG = False Then
'                         SELRREQ_QR(0) = txt_dept.Text
'                         SELRREQ_QR(1) = txt_modcode.Text
'                         SELRREQ_QR(2) = Msk_caseno.FormattedText
'                         Set SELRREQ_RS = SELRREQ_QR.OpenResultset(1, 3)
'                         SELRREQFLAG = True
'                    Else
'                         SELRREQ_QR(0) = txt_dept.Text
'                         SELRREQ_QR(1) = txt_modcode.Text
'                         SELRREQ_QR(2) = Msk_caseno.FormattedText
'                         SELRREQ_RS.MoreResults
'                         SELRREQ_RS.Requery
'
'                     End If
'                     If SELRREQ_RS.RowCount > 0 Then
'                         cmb_preinvest.clear
'                         While Not SELRREQ_RS.EOF
'                             cmb_preinvest.AddItem SELRREQ_RS("REQNNO")
'                             SELRREQ_RS.MoveNext
'                         Wend
'                         cmb_preinvest.Enabled = True
'                     End If
''commented by mono
'                    If seLdreqrsflag = False Then
'                        SELDREQQR(0) = Msk_caseno.FormattedText
'                        Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
'                        seLdreqrsflag = True
'                    Else
'                        SELDREQQR(0) = Msk_caseno.FormattedText
'                        SELDREQRS.MoreResults
'                        SELDREQRS.Requery
'                    End If
'                If SELDREQRS.RowCount > 0 Then
'                    cmb_preinvest.clear
'                    While Not SELDREQRS.EOF
'                        cmb_preinvest.AddItem SELDREQRS("PATHNO") & "$" & SELDREQRS("REQNO")
'                        SELDREQRS.MoveNext
'                    Wend
'                    cmb_preinvest.ListIndex = 0
'                End If
                    ' Previously Investigated requisition nos
                    
                    
                    
                End If
          End If
''          If txt_modcode.Text = "IH" Then
''                immunohistochemistry.Visible = True
''                cmb_ihprocode.Text = SELRREQUISITN_RS2("SER_CODE")
''
''                If SELMA101FLAG = False Then
''                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS2("SER_CODE"))
''                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
''                    SELMA101FLAG = True
''                Else
''                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS2("SER_CODE"))
''                    SELMA101_RS.MoreResults
''                    SELMA101_RS.Requery
''                End If
''                If SELMA101_RS.RowCount > 0 Then
''                    cmb_ihprocedure.Text = SELMA101_RS("SER_DET")
''                End If
''                If SELDREQUISITNFLAG = False Then
''                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
''                    Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
''                    SELDREQUISITNFLAG = True
''                Else
''                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
''                    seldrequisitn_rs.MoreResults
''                    seldrequisitn_rs.Requery
''                End If
''                cmb_Ihsoutsideblock = seldrequisitn_rs("OUTBLOCK")
''                cmb_Ihsoutslide.Text = seldrequisitn_rs("OUTSLIDES")
''                If Len(Trim(seldrequisitn_rs("DURILLNESS"))) > 1 Then
''
''                    i = InStr(seldrequisitn_rs("DURILLNESS"), "/")
''                    txt_year.Text = Mid(seldrequisitn_rs("DURILLNESS"), 1, i - 1)
''                    CMB_DURILLMONTH.Text = Mid(seldrequisitn_rs("DURILLNESS"), i + 1, 2)
''                End If
''                'cmb_Ihspreinv.Text = Trim(SELRREQUISITN_RS2("PREVEXAM"))
''                cmb_refdoc.Text = Trim(seldrequisitn_rs("ReFDOC"))
''                cmb_reftodoc.Text = Trim(seldrequisitn_rs("REFTODOC"))
''                If seldrequisitn_rs.EOF = False Then
''                    txt_ihmeterial.Text = " " & seldrequisitn_rs("NMETERIALS").Value
''                Else
''                    txt_ihmeterial.Text = ""
''                End If
''                IHENABL ("DIS")
''                lbl_ihreqno.Visible = False
''                lbl_ihpathno.Visible = False
''                txt_ihreqno.Visible = False
''                txt_ihpathno.Visible = False
''                ' Previously Investigated requisition nos
''                    If SELRREQFLAG = False Then
''                         SELRREQ_QR(0) = txt_dept.Text
''                         SELRREQ_QR(1) = txt_modcode.Text
''                         SELRREQ_QR(2) = Msk_caseno.FormattedText
''                         Set SELRREQ_RS = SELRREQ_QR.OpenResultset(1, 3)
''                         SELRREQFLAG = True
''                    Else
''                         SELRREQ_QR(0) = txt_dept.Text
''                         SELRREQ_QR(1) = txt_modcode.Text
''                         SELRREQ_QR(2) = Msk_caseno.FormattedText
''                         SELRREQ_RS.MoreResults
''                         SELRREQ_RS.Requery
''
''                     End If
''                     If SELRREQ_RS.RowCount > 0 Then
''                         cmb_preinvest.clear
''                         While Not SELRREQ_RS.EOF
''                             cmb_Ihspreinv.AddItem SELRREQ_RS("REQNNO")
''                             SELRREQ_RS.MoveNext
''                         Wend
''                         cmb_Ihspreinv.Enabled = True
''
''                     End If
''                     If SELRREQUISITN_RS2.RowCount > 0 Then
''                                If SELRREQUISITN_RS2("RECORDNO") = 1 Then
''                                CMB_IHNIP.Text = "VPP"
''                                End If
''                                If SELRREQUISITN_RS2("RECORDNO") = 2 Then
''                                CMB_IHNIP.Text = "LR"
''                                End If
''                                If SELRREQUISITN_RS2("RECORDNO") = 3 Then
''                                CMB_IHNIP.Text = "BAR"
''                                End If
''                                If SELRREQUISITN_RS2("RECORDNO") = 4 Then
''                                CMB_IHNIP.Text = "DrOp"
''                                End If
''                            End If
''                    ' Previously Investigated requisition nos
''
''
''
''
''          End If
''          '**** FA FILLING
''          If Trim(txt_modcode.Text) = "FA" Then
''                 fnac.Visible = True
''                FAENABL ("DIS")
''                If Left(Msk_caseno.FormattedText, 2) = "RF" Then
''                        cmb_outsideslidesfn.Visible = True
''                        lbl_outsideslides.Visible = True
''                Else
''                        cmb_outsideslidesfn.Visible = False
''                        lbl_outsideslides.Visible = False
''                End If
''                txt_FNreqno.Visible = False
''                lbl_fnreqno.Visible = False
''                cmb_fnprocode.Text = SELRREQUISITN_RS2("SER_CODE")
''
''
''                If SELMA101FLAG = False Then
''                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS2("SER_CODE"))
''                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
''                    SELMA101FLAG = True
''                Else
''                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS2("SER_CODE"))
''                    SELMA101_RS.MoreResults
''                    SELMA101_RS.Requery
''                End If
''                If SELMA101_RS.RowCount > 0 Then
''                    CMB_FNPROCEDURE.Text = SELMA101_RS("SER_DET")
''                End If
''                If SELRREQUISITN_RS2.RowCount > 0 Then
''                    If SELRREQUISITN_RS2("RECORDNO") = 1 Then
''                    CMB_FANIP.Text = "VPP"
''                    End If
''                    If SELRREQUISITN_RS2("RECORDNO") = 2 Then
''                    CMB_FANIP.Text = "LR"
''                    End If
''                    If SELRREQUISITN_RS2("RECORDNO") = 3 Then
''                    CMB_FANIP.Text = "BAR"
''                    End If
''                    If SELRREQUISITN_RS2("RECORDNO") = 4 Then
''                    CMB_FANIP.Text = "DrOp"
''                    End If
''                End If
''                If SELDREQUISITNFLAG = False Then
''                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
''                    Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
''                    SELDREQUISITNFLAG = True
''                Else
''                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
''                    seldrequisitn_rs.MoreResults
''                    seldrequisitn_rs.Requery
''                End If
''                If seldrequisitn_rs.RowCount > 0 Then
''                    If SELMA301FLAG = False Then
''                        SELMA301QR(0) = seldrequisitn_rs("SITE")
''                        Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
''                        SELMA301FLAG = True
''                    Else
''                        SELMA301QR(0) = seldrequisitn_rs("SITE")
''                        Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
''                        SELMA301RS.MoreResults
''                        SELMA301RS.Requery
''
''                    End If
''                    If SELMA301RS.RowCount > 0 Then
''                        If Trim(SELMA301RS("SITEDESC")) = "site not specified/recorded" Then
''                            cmb_fnsite.Text = ""
''                        Else
''                            cmb_fnsite.Text = SELMA301RS("SITEDESC")
''                        End If
''                    Else
''                            cmb_fnsite.Text = ""
''                    End If
''                    cmb_refdoc.Text = Trim(seldrequisitn_rs("REFDOC"))
''                    cmb_reftodoc.Text = Trim(seldrequisitn_rs("reftodoc"))
''                    i = InStr(seldrequisitn_rs("durillness"), "/")
''                    If i <> 0 Then
''                    txt_year.Text = Mid(Trim(seldrequisitn_rs("durillness")), 1, i - 1)
''                    CMB_DURILLMONTH.Text = Mid(Trim(seldrequisitn_rs("durillness")), i + 1, 2)
''                    End If
''                    cmb_fnlaterality.Text = Trim(seldrequisitn_rs("LATERALITS"))
''                    txt_fnmetastasis.Text = Trim(seldrequisitn_rs("METASTASIS"))
''                    cmb_fnlaterality1.Text = Trim(seldrequisitn_rs("LATERALITM"))
''                    TXT_FNSLIDEDRY.Text = Trim(seldrequisitn_rs("OUTSSPECIM"))
''                    TXT_FNOSLIDEWET.Text = Trim(seldrequisitn_rs("OUTBLOCK"))
''                    cmb_outsideslidesfn.Text = Trim(seldrequisitn_rs("OUTSLIDES"))
''                    If seldrequisitn_rs.EOF = False Then
''                        txt_fnnmeterial.Text = " " & seldrequisitn_rs("NMETERIALS").Value
''                    Else
''                        txt_fnnmeterial.Text = ""
''                    End If
''                    cmb_fnac.Text = Trim(seldrequisitn_rs("OPPROCEDUR"))
''                    If Trim(seldrequisitn_rs("frozensec")) = "" Then
''                    Else
'''                    CMB_CELLBLOCK.AddItem Trim(seldrequisitn_rs("FROZENSEC"))
'''                    CMB_CELLBLOCK.Text = Trim(seldrequisitn_rs("FROZENSEC"))
''                    End If
''                End If
''          End If
''   '***** CG OR CN
''           If Trim(txt_modcode.Text) = "CG" Or Trim(txt_modcode.Text) = "CN" Then
''                CYTOLOGY.Visible = True
''                CGENABL ("DIS")
''                txt_cyreqno1.Visible = False
''                lbl_cyreqno.Visible = False
''
''                cbo_proccode.Text = SELRREQUISITN_RS2("SER_CODE")
''
''
''                If SELMA101FLAG = False Then
''                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS2("SER_CODE"))
''                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
''                    SELMA101FLAG = True
''                Else
''                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS2("SER_CODE"))
''                    SELMA101_RS.MoreResults
''                    SELMA101_RS.Requery
''                End If
''                If SELMA101_RS.RowCount > 0 Then
''                    CMB_CGPROCEDURE.Text = SELMA101_RS("SER_DET")
''                End If
''                If SELRREQUISITN_RS2.RowCount > 0 Then
''                    If SELRREQUISITN_RS2("RECORDNO") = 1 Then
''                    CMB_CGNIP.Text = "VPP"
''                    End If
''                    If SELRREQUISITN_RS2("RECORDNO") = 2 Then
''                    CMB_CGNIP.Text = "LR"
''                    End If
''                    If SELRREQUISITN_RS2("RECORDNO") = 3 Then
''                    CMB_CGNIP.Text = "BAR"
''                    End If
''                    If SELRREQUISITN_RS2("RECORDNO") = 4 Then
''                    CMB_CGNIP.Text = "DrOp"
''                    End If
''                End If
''                If SELDREQUISITNFLAG = False Then
''                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
''                    Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
''                    SELDREQUISITNFLAG = True
''                Else
''                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
''                    seldrequisitn_rs.MoreResults
''                    seldrequisitn_rs.Requery
''                End If
''                If seldrequisitn_rs.RowCount > 0 Then
''                    If SELMA301FLAG = False Then
''                        SELMA301QR(0) = seldrequisitn_rs("SITE")
''                        Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
''                        SELMA301FLAG = True
''                    Else
''                        SELMA301QR(0) = seldrequisitn_rs("SITE")
''                        Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
''                        SELMA301RS.MoreResults
''                        SELMA301RS.Requery
''
''                    End If
''                    If SELMA301RS.RowCount > 0 Then
''                        If Trim(SELMA301RS("SITEDESC")) = "site not specified/recorded" Then
''                            cmb_cnsite.Text = ""
''                        Else
''                            cmb_cnsite.Text = SELMA301RS("SITEDESC")
''                        End If
''                    Else
''                            cmb_cnsite.Text = ""
''                    End If
''                    cmb_refdoc.Text = Trim(seldrequisitn_rs("REFDOC"))
''                    cmb_claterality.Text = Trim(seldrequisitn_rs("LATERALITS"))
''                    txt_cmetastasis.Text = Trim(seldrequisitn_rs("METASTASIS"))
''
''                    'cmb_cyoutspecimen.Text = Trim(seldrequisitn_rs("OUTSSPECIM"))
''                    cmb_cnoslides.Text = Trim(seldrequisitn_rs("OUTBLOCK"))
''                    cmb_coutslides.Text = Trim(seldrequisitn_rs("OUTSLIDES"))
''                    If seldrequisitn_rs.EOF = False Then
''                        txt_cnatmeterial.Text = " " & Trim(seldrequisitn_rs("NMETERIALS"))
''                    Else
''                        txt_cnatmeterial.Text = " "
''                    End If
''
''                    If txt_modcode.Text = "CG" Then
''                        lbl_lmp.Visible = True
''                        MSK_LMP.Visible = True
''                        If IsNull(Trim(seldrequisitn_rs("FROZENDATE"))) Then
''                        Else
''                        MSK_LMP.Text = Trim(seldrequisitn_rs("FROZENDATE"))
''                        End If
''                    Else
''                        lbl_lmp.Visible = False
''                        MSK_LMP.Visible = False
''                    End If
''                    If seldsftdcyflag = False Then
''                        seldsftdcyqr(0) = req_chglen(cmb_reqno.Text)
''                        Set seldsftdcyrs = seldsftdcyqr.OpenResultset(1, 3)
''                        seldsftdcyflag = True
''                    Else
''                        seldsftdcyqr(0) = req_chglen(cmb_reqno.Text)
''                        seldsftdcyrs.MoreResults
''                        seldsftdcyrs.Requery
''                    End If
''                    If seldsftdcyrs.RowCount > 0 Then
''
''                        If seldsftdcyrs.EOF = False Then
''                            txt_scopyfind.Text = seldsftdcyrs("spfind")
''                         Else
''                            txt_scopyfind.Text = ""
''                        End If
''                        If seldsftdcyrs.EOF = False Then
''                            txt_trtdetail.Text = " " & seldsftdcyrs("trdet")
''                         Else
''                            txt_trtdetail.Text = ""
''                        End If
''                        If seldsftdcyrs.EOF = False Then
''                            cmb_cyoutspecimen.Text = seldsftdcyrs("outspec")
''                         Else
''                            cmb_cyoutspecimen.Text = ""
''                        End If
''                    End If
''                End If
''          End If
''
''   '***** CG OR CN
''   '***** BIOCHEMISTRY
''           If Trim(txt_modcode.Text) = "BC" Then
''                If SELDREQUISITNFLAG = False Then
''                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
''                    Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
''                    SELDREQUISITNFLAG = True
''                Else
''                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
''                    seldrequisitn_rs.MoreResults
''                    seldrequisitn_rs.Requery
''                End If
''                If seldrequisitn_rs.RowCount > 0 Then
''                    biochemistry.Msk_caseno.Text = Msk_caseno.FormattedText
''                    biochemistry.CMB_BODYF.Text = Trim(seldrequisitn_rs("METASTASIS"))
''                    biochemistry.txt_name.Text = txt_name
''                End If
''                biochemistry.show
''                If SELDMULPROCEFLAG = False Then
''                    SELDMULPROCEQR(0) = req_chglen(cmb_reqno.Text)
''                    Set SELDMULPROCERS = SELDMULPROCEQR.OpenResultset(1, 3)
''                    SELDMULPROCEFLAG = True
''               Else
''                    SELDMULPROCEQR(0) = req_chglen(cmb_reqno.Text)
''                    SELDMULPROCERS.MoreResults
''                    SELDMULPROCERS.Requery
''                End If
''                If SELDMULPROCERS.RowCount > 0 Then
''                        While Not SELDMULPROCERS.EOF
''                            BCVIEW
''                            SELDMULPROCERS.MoveNext
''                        Wend
''                        BCFLAG = True
''                End If
''           End If
''           If Trim(txt_modcode.Text) = "EL" Then
''                If SELDREQUISITNFLAG = False Then
''                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
''                    Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
''                    SELDREQUISITNFLAG = True
''                Else
''                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
''                    seldrequisitn_rs.MoreResults
''                    seldrequisitn_rs.Requery
''                End If
''                If seldrequisitn_rs.RowCount > 0 Then
''                  Em_lab.MaskEdBox1 = Msk_caseno.FormattedText
''
''                    Em_lab.txt_name.Text = txt_name
''                End If
''                Em_lab.show
''                If SELDMULPROCEFLAG = False Then
''                    SELDMULPROCEQR(0) = req_chglen(cmb_reqno.Text)
''                    Set SELDMULPROCERS = SELDMULPROCEQR.OpenResultset(1, 3)
''                    SELDMULPROCEFLAG = True
''               Else
''                    SELDMULPROCEQR(0) = req_chglen(cmb_reqno.Text)
''                    SELDMULPROCERS.MoreResults
''                    SELDMULPROCERS.Requery
''                End If
''                If SELDMULPROCERS.RowCount > 0 Then
''                    While Not SELDMULPROCERS.EOF
''                        ELVIEW
''                        SELDMULPROCERS.MoveNext
''                    Wend
''                    ELFLAG = True
''                End If
''           End If
''           If Trim(txt_modcode.Text) = "TM" Then
''                If SELDREQUISITNFLAG = False Then
''                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
''                    Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
''                    SELDREQUISITNFLAG = True
''                Else
''                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
''                    seldrequisitn_rs.MoreResults
''                    seldrequisitn_rs.Requery
''                End If
''                If seldrequisitn_rs.RowCount > 0 Then
''                    TUMOURMARKER.txt_clininfo = txt_clinhist
''                    TUMOURMARKER.txt_Dosage = seldrequisitn_rs("PREPATHNO")
''                    If seldrequisitn_rs.EOF = False Then
''                        TUMOURMARKER.cmb_nmeterial = " " & seldrequisitn_rs("NMETERIALS")
''                    Else
''                        TUMOURMARKER.cmb_nmeterial = ""
''                    End If
''                    TUMOURMARKER.txt_rem = seldrequisitn_rs("METASTASIS")
''                End If
''                TUMOURMARKER.show
''                If SELDMULPROCEFLAG = False Then
''                    SELDMULPROCEQR(0) = req_chglen(cmb_reqno.Text)
''                    Set SELDMULPROCERS = SELDMULPROCEQR.OpenResultset(1, 3)
''                    SELDMULPROCEFLAG = True
''               Else
''                    SELDMULPROCEQR(0) = req_chglen(cmb_reqno.Text)
''                    SELDMULPROCERS.MoreResults
''                    SELDMULPROCERS.Requery
''                End If
''                If SELDMULPROCERS.RowCount > 0 Then
''
''                    While Not SELDMULPROCERS.EOF
''                        TMVIEW
''                        SELDMULPROCERS.MoveNext
''                    Wend
''
''                    TUMOURMARKER.CHKCLEAR
''                    tmflag = True
''                End If
''           End If
''   '***** BIOCHEMISTRY
''   '***** MICROBIOLOGY
''        If Trim(txt_modcode.Text) = "BA" Or Trim(txt_modcode.Text) = "MB" Or Trim(txt_modcode.Text) = "SE" Or Trim(txt_modcode.Text) = "MC" Or Trim(txt_modcode.Text) = "CM" Or Trim(txt_modcode.Text) = "MO" Or Trim(txt_modcode.Text) = "XX" Then
''                Microbiology.Visible = True
''                MBENABL ("DIS")
''                MBCLEAR
''                txt_mbreqno.Visible = False
''                lbl_mbreqno.Visible = False
''                TXT_MBPATHNO.Visible = False
''                LBL_MBMICRONO.Visible = False
''                txt_pathno.Visible = True
''                lbl_pathno.Visible = True
''                I1 = 0
''                If SELDMULPROCEFLAG = False Then
''                    SELDMULPROCEQR(0) = req_chglen(cmb_reqno.Text)
''                    Set SELDMULPROCERS = SELDMULPROCEQR.OpenResultset(1, 3)
''                    SELDMULPROCEFLAG = True
''               Else
''                    SELDMULPROCEQR(0) = req_chglen(cmb_reqno.Text)
''                    SELDMULPROCERS.MoreResults
''                    SELDMULPROCERS.Requery
''                End If
''                If SELDMULPROCERS.RowCount > 0 Then
''                    List_code.clear
''                    list_proc.clear
''                    While Not SELDMULPROCERS.EOF
''                        List_code.AddItem SELDMULPROCERS("PRO_CODE")
''
''                        If SELMA101FLAG = False Then
''                            SELMA101_QR(0) = Trim(SELDMULPROCERS("pro_CODE"))
''                            Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
''                            SELMA101FLAG = True
''                        Else
''                            SELMA101_QR(0) = Trim(SELDMULPROCERS("pro_CODE"))
''                            SELMA101_RS.MoreResults
''                            SELMA101_RS.Requery
''                        End If
''                        If SELMA101_RS.RowCount > 0 Then
''                            While Not SELMA101_RS.EOF
''                                list_proc.AddItem SELMA101_RS("SER_DET")
''
''                                SELMA101_RS.MoveNext
''                            Wend
''                         End If
''                         SELDMULPROCERS.MoveNext
''                    Wend
''                        For i = 0 To List_code.ListCount - 1
''                            List_code.Selected(i) = True
''
''                        Next i
''                        For i = 0 To list_proc.ListCount - 1
''                            list_proc.Selected(i) = True
''                        Next i
''
''                End If
''
''
''
''                If SELDREQUISITNFLAG = False Then
''                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
''                    Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
''                    SELDREQUISITNFLAG = True
''                Else
''                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
''                    seldrequisitn_rs.MoreResults
''                    seldrequisitn_rs.Requery
''                End If
''
''                    cmb_refdoc.Text = Trim(seldrequisitn_rs("REFDOC"))
''                    cmb_reftodoc.Text = Trim(seldrequisitn_rs("reftodoc"))
''                    lbl_pathno.Visible = True
''                    txt_pathno.Visible = True
''                    txt_pathno.Text = Trim(seldrequisitn_rs("pathno"))
''                    i = InStr(seldrequisitn_rs("durillness"), "/")
''                    If i <> 0 Then
''                    txt_year.Text = Mid(Trim(seldrequisitn_rs("durillness")), 1, i - 1)
''                    CMB_DURILLMONTH.Text = Mid(Trim(seldrequisitn_rs("durillness")), i + 1, 2)
''                    End If
''                    If seldrequisitn_rs.EOF = False Then
''                        cmb_mboutslide.Text = " " & (seldrequisitn_rs("OUTSLIDES"))
''                    Else
''                        cmb_mboutslide.Text = ""
''                    End If
''                    If SELSPONANTIBFLAG = False Then
''                        SELSPONANTIBQR(0) = req_chglen(cmb_reqno.Text)
''                        Set SELSPONANTIBRS = SELSPONANTIBQR.OpenResultset(1, 3)
''                        SELSPONANTIBFLAG = True
''                    Else
''                        SELSPONANTIBQR(0) = req_chglen(cmb_reqno.Text)
''                        SELSPONANTIBRS.MoreResults
''                        SELSPONANTIBRS.Requery
''                    End If
''                    If SELSPONANTIBRS.RowCount > 0 Then
''                        If SELSPONANTIBRS.EOF = False Then
''                            txt_mbspecimen.Text = Trim(SELSPONANTIBRS("SPECIMEN").Value)
''                        Else
''                            txt_mbspecimen.Text = ""
''                        End If
''                        If SELSPONANTIBRS.EOF = False Then
''                            txt_ongoinganti.Text = Trim(SELSPONANTIBRS("ONGOINGANT").Value)
''                        Else
''                            txt_ongoinganti.Text = ""
''                        End If
''                    End If
''          End If
''
''   '***** MICROBIOLOGY
       '**** FA FILLING
          If Trim(txt_modcode.Text) = "FA" Then
                FNACVIEW
          End If
   '***** CG OR CN
           If Trim(txt_modcode.Text) = "CG" Or Trim(txt_modcode.Text) = "CN" Or Trim(txt_modcode.Text) = "PG" Or Trim(txt_modcode.Text) = "PN" Then
                CGVIEW
           End If
   '***** CG OR CN
   '***** BIOCHEMISTRY
           If Trim(txt_modcode.Text) = "BC" Then
                BIOVIEW
           End If
           If Trim(txt_modcode.Text) = "EL" Then
                EMERVIEW
           End If
           If Trim(txt_modcode.Text) = "TM" Then
                TUMOURVIEW
           End If
   '***** BIOCHEMISTRY
   '***** MICROBIOLOGY
        If Trim(txt_modcode.Text) = "BA" Or Trim(txt_modcode.Text) = "MB" Or Trim(txt_modcode.Text) = "SE" Or Trim(txt_modcode.Text) = "MC" Or Trim(txt_modcode.Text) = "CM" Or Trim(txt_modcode.Text) = "MO" Or Trim(txt_modcode.Text) = "XX" Then
                MBVIEW
        End If
   '***** MICROBIOLOGY
   
   '***** ENDOSCOPY
            If Trim(txt_modcode.Text) = "ED" Then
                ENDOSVIEW
            End If
   
   '***** ENDOSCOPY
   '***** CARDIOLOGY
           If Trim(txt_modcode.Text) = "2D" Or Trim(txt_modcode.Text) = "EK" Or Trim(txt_modcode.Text) = "PF" Or Trim(txt_modcode.Text) = "ST" Then
                CARDVIEW
           End If
   '***** CARDIOLOGY
   '***** HAEMATOLOGY
           If Trim(txt_modcode.Text) = "HL" Then
                HEMVIEW
           End If
   '***** HAEMATOLOGY
   '***** BM
        If Trim(txt_modcode.Text) = "BM" Then
                BONEMVIEW
        End If
   '***** BM
   '**** BF
          If Trim(txt_modcode.Text) = "BF" Then
                BODYFVIEW
          End If
   '**** BF
   '***** HLA TYPING
           If Trim(txt_modcode.Text) = "HT" Then
                HLAVIEW
                
           End If
   '***** HLA TYPING
   '***** DTM LAB
           If Trim(txt_modcode.Text) = "SS" Then
                dtmview
           End If
   '***** DTM LAB
   '***** BLOOD PRODUCTS
          If Trim(txt_modcode.Text) = "BP" Then
               BPVIEW
          End If
   '***** BLOOD PRODUCTS
          
    
         '*****
    End If
End If
'If Trim(txt_catgory.Text) = "NC" Then
          CMB_MEMO.Enabled = False
'End If
Exit Sub
MSG:
MsgBox Err.Description

End Sub

Public Sub SPENABLE(EN As String)
If EN = "DIS" Then
        cmb_procedure.Enabled = False
        cmb_procedurecode.Enabled = False
        cmb_sites.Enabled = False
        cmb_laterality1.Enabled = False
        txt_metastasis.Enabled = False
        txt_outspecimen.Enabled = False
        cmb_outblock.Enabled = False
        cmb_outslide.Enabled = False
        cmb_frozensec.Enabled = False
        MSK_FROSECDT.Enabled = False
        cmb_laterality.Enabled = False
        TXT_OPPROCEDURE.Enabled = False
        cmb_preinvest.Enabled = False
        lbl_spreqno1.Visible = True
        lbl_sppathno.Visible = False
        txt_sppathno.Visible = True
        txt_spreqno.Visible = False
        txt_nmeterial.Enabled = False
        cmb_vpp.Enabled = False
        cmb_spfnac.Enabled = False
'        lbl_pathno.Visible = True
'        txt_pathno.Visible = True
'        lbl_pathno.Enabled = True
'        txt_pathno.Enabled = False
End If
If EN = "ENA" Then
        cmb_procedure.Enabled = True
        cmb_procedurecode.Enabled = True
        cmb_sites.Enabled = True
        cmb_laterality1.Enabled = True
        txt_metastasis.Enabled = True
        txt_outspecimen.Enabled = True
        cmb_outblock.Enabled = True
        cmb_outslide.Enabled = True
        cmb_frozensec.Enabled = True
        MSK_FROSECDT.Enabled = True
        cmb_laterality.Enabled = True
        TXT_OPPROCEDURE.Enabled = True
        cmb_preinvest.Enabled = True
        lbl_spreqno1.Visible = True
        lbl_sppathno.Visible = True
        txt_sppathno.Visible = False
        txt_spreqno.Visible = True
        txt_nmeterial.Enabled = True
        cmb_vpp.Enabled = True
        cmb_spfnac.Enabled = True
'        lbl_pathno.Visible = True
'        txt_pathno.Visible = True
'        lbl_pathno.Enabled = True
'        txt_pathno.Enabled = False
End If
If EN = "MOD" Then
        cmb_procedure.Enabled = False
        cmb_procedurecode.Enabled = False
        cmb_sites.Enabled = True
        cmb_laterality1.Enabled = True
        txt_metastasis.Enabled = True
        txt_outspecimen.Enabled = True
        cmb_outblock.Enabled = True
        cmb_outslide.Enabled = True
        cmb_frozensec.Enabled = True
        MSK_FROSECDT.Enabled = True
        cmb_laterality.Enabled = True
        TXT_OPPROCEDURE.Enabled = True
        cmb_preinvest.Enabled = True
        lbl_spreqno1.Visible = False
        lbl_sppathno.Visible = True
        txt_sppathno.Visible = False
        txt_spreqno.Visible = False
        txt_nmeterial.Enabled = True
        cmb_refdoc.Enabled = True
        cmb_reftodoc.Enabled = True
        CMB_DURILLMONTH.Enabled = True
        txt_year.Enabled = True
        cmb_vpp.Enabled = True
        txt_indctinvstg.Enabled = True
        txt_clinhist.Enabled = True
        cmb_spfnac.Enabled = True
'        lbl_pathno.Visible = True
'        txt_pathno.Visible = True
'        lbl_pathno.Enabled = True
'        txt_pathno.Enabled = False
End If
End Sub

Public Sub RCLEAR()
  Msk_caseno.Mask = "??>/######"
  Msk_caseno.Text = ""
  txt_name.Text = ""
  txt_sex.Text = ""
  txt_Age.Text = ""
  txt_catgory.Text = ""
  txt_ward.Text = ""
  txt_status.Text = ""
  
  txt_indctinvstg.Text = ""
  cmb_dept.Text = ""
  cmb_modality.Text = ""
  txt_modcode.Text = ""
  txt_dept.Text = ""
  cmb_refdoc.Text = ""
  cmb_reftodoc.Text = ""
  txt_pridocser.Text = ""
  cmb_reqno.Text = ""
  txt_pathno.Text = ""
  txt_bedno.Text = ""
  txt_year.Text = ""
  CMB_DURILLMONTH.Text = ""
  txt_clinhist.Text = ""
  cmb_ihprocedure.Text = " "
  cmb_procedurecode.Text = " "
  TEXT1.Text = ""
  
End Sub

Public Sub SPCLEAR()
cmb_procedure.Text = ""
cmb_procedurecode.Text = ""
cmb_sites.Text = ""
cmb_laterality.Text = ""
cmb_laterality1.Text = ""
txt_metastasis.Text = ""
'TXT_OPPROCEDURE.Text = ""
txt_outspecimen.Text = ""
cmb_outblock.Text = ""
cmb_outslide.Text = ""
cmb_frozensec.Text = ""
MSK_FROSECDT.Mask = "##/##/####"
MSK_FROSECDT.Text = "__/__/____"
cmb_preinvest.Text = ""
txt_nmeterial.Text = ""
txt_sppathno.Text = ""
cmb_spfnac.Text = ""
cmb_vpp.Text = ""
End Sub

Public Sub IH()
On Error GoTo MSG:
If txt_modcode.Text = "IH" Then
        
        surgicalPathology.Visible = False
        immunohistochemistry.Visible = True
        frm_hematopathology.Visible = False
        biochemistry.Visible = False
        bceltm.Visible = False
'        cardiology.Visible = False
        Microbiology.Visible = False
        CYTOLOGY.Visible = False
        
        fnac.Visible = False
        'cpl.Visible = False
        TUMOURMARKER.Visible = False
        hemcbc.Hide
        HEMBF.Hide
        HEMBMA.Hide
        'HLATYPING.Visible = False
        lbl_req.Left = 3000
        lbl_req.Caption = "REQUISITION - IMMUNOHISTOCHEMISTRY"
        pathology.Caption = "Pathology - Immunohistochemistry"
        'txt_dept.Text = "FDZ"
        
      ' PROCEDURE CODE FILLING IN THE COMBOBOX
    If SELRINVESTFLAG = False Then
        SELRINVEST_QR(0) = Trim(txt_modcode.Text)
        Set SELRINVEST_RS = SELRINVEST_QR.OpenResultset(1, 3)
        SELRINVESTFLAG = True
    Else
        SELRINVEST_QR(0) = Trim(txt_modcode.Text)
        SELRINVEST_RS.MoreResults
        SELRINVEST_RS.Requery
    End If
    If SELRINVEST_RS.RowCount > 0 Then
        cmb_procedure.clear
        cmb_procedurecode.clear
        While Not SELRINVEST_RS.EOF
            cmb_ihprocode.AddItem Trim(SELRINVEST_RS("ser_code"))
            If SELMA101FLAG = False Then
                SELMA101_QR(0) = Trim(SELRINVEST_RS("SER_CODE"))
                Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                SELMA101FLAG = True
             Else
                SELMA101_QR(0) = SELRINVEST_RS("SER_CODE")
                SELMA101_RS.MoreResults
                SELMA101_RS.Requery
            End If
            If SELMA101_RS.RowCount > 0 Then
                'cmb_ihprocedure.clear
                While Not SELMA101_RS.EOF
                cmb_ihprocedure.AddItem SELMA101_RS("SER_DET")
                SELMA101_RS.MoveNext
                Wend
            End If
            SELRINVEST_RS.MoveNext
        Wend
    End If
    
    ' PROCEDURE CODE FILLING IN THE COMBOBOX
    
    
    If IHFLAG = False Then
        txt_sppathreqno.Visible = False
        lbl_spreqno.Visible = False
        cmb_Ihsoutsideblock.Visible = True
        cmb_Ihsoutslide.Visible = True
        lbl_hisblock.Visible = True
       lbl_hisslide.Visible = True
        cmb_Ihsoutsideblock.clear
        cmb_Ihsoutslide.clear
        For I = 0 To 5
            cmb_Ihsoutsideblock.AddItem I
            cmb_Ihsoutslide.AddItem I
        Next I
    Else
        txt_sppathreqno.Visible = True
        lbl_spreqno.Visible = True
        cmb_Ihsoutsideblock.Visible = False
        cmb_Ihsoutslide.Visible = False
       lbl_hisblock.Visible = False
       lbl_hisslide.Visible = False
    End If
    IHENABL ("ENA")
    ' FILLING THE PREVIOUSLY INVESTIGATED REQ NO.
   If SELRREQFLAG = False Then
        SELRREQ_QR(0) = txt_dept.Text
        SELRREQ_QR(1) = txt_modcode.Text
        SELRREQ_QR(2) = Msk_caseno.FormattedText
        Set SELRREQ_RS = SELRREQ_QR.OpenResultset(1, 3)
        SELRREQFLAG = True
   Else
        SELRREQ_QR(0) = txt_dept.Text
        SELRREQ_QR(1) = txt_modcode.Text
        SELRREQ_QR(2) = Msk_caseno.FormattedText
        SELRREQ_RS.MoreResults
        SELRREQ_RS.Requery
   
    End If
    If SELRREQ_RS.RowCount > 0 Then
        cmb_preinvest.clear
        While Not SELRREQ_RS.EOF
            cmb_Ihspreinv.AddItem SELRREQ_RS("REQNNO")
            SELRREQ_RS.MoveNext
        Wend
        cmb_Ihspreinv.ListIndex = 0
    End If
  ' FILLING THE PREVIOUSLY INVESTIGATED REQ NO.
    If REPEATFLAG = True Then
        lbl_ihpathno.Visible = True
        lbl_ihreqno.Visible = True
        txt_ihpathno.Visible = True
        txt_ihreqno.Visible = True
        txt_ihpathno.Enabled = False
        txt_ihreqno.Enabled = False
        REPEATFLAG = False
    End If
    lbl_ihpathno.Visible = True
    txt_ihpathno.Visible = True
 End If
 Exit Sub
MSG:
 MsgBox Err.Description
 
End Sub

Public Sub IHASSIGN()
Dim durill As String
On Error GoTo MSG:
'*** INSERTING INT0 RREQUISITION
INSRREQUISITN_QR(0) = req_chglen(Trim(txt_ihreqno.Text))
INSRREQUISITN_QR(1) = GET_REVDATE(CStr(DATE11))
INSRREQUISITN_QR(2) = Trim(txt_time.Text)
INSRREQUISITN_QR(3) = Trim(Msk_caseno.FormattedText)
INSRREQUISITN_QR(4) = Trim(txt_dept.Text)
INSRREQUISITN_QR(5) = Trim(txt_modcode.Text)
INSRREQUISITN_QR(6) = Trim(txt_status.Text)
INSRREQUISITN_QR(7) = Trim(txt_catgory.Text)
INSRREQUISITN_QR(8) = Trim(cmb_ihprocode.Text)
INSRREQUISITN_QR(9) = Trim(txt_pridocser.Text)
    If SELRREQUISITNFLAG = False Then
       SELRREQUISITN_QR(0) = Msk_caseno.FormattedText
       Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
       SELRREQUISITNFLAG = True
    Else
        SELRREQUISITN_QR(0) = Msk_caseno.FormattedText
        SELRREQUISITN_RS.MoreResults
        SELRREQUISITN_RS.Requery
    
    End If
    If SELRREQUISITN_RS.RowCount > 0 Then
        INSRREQUISITN_QR(10) = "YES"
    Else
        INSRREQUISITN_QR(10) = "NO"
    End If
NO_CHARGEIH
If No_Chargeflag = True Then
    INSRREQUISITN_QR(11) = "NOCHARGE"
    INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
    INSRREQUISITN_QR(13) = "P"
    No_Chargeflag = False
Else
    If Trim(CMB_IHNIP.Text) = "LR" Or Trim(CMB_IHNIP.Text) = "BAR" Or Trim(CMB_IHNIP.Text) = "DrOp" Then
            
        If Trim(CMB_IHNIP.Text) = "LR" Then
            INSRREQUISITN_QR(11) = "LR"
            INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
            INSRREQUISITN_QR(13) = "P"
        End If
        If Trim(CMB_IHNIP.Text) = "BAR" Then
            INSRREQUISITN_QR(11) = "BAR"
            INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
            INSRREQUISITN_QR(13) = "P"
        End If
        If Trim(CMB_IHNIP.Text) = "DrOp" Then
            INSRREQUISITN_QR(11) = "DrOp"
            INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
            INSRREQUISITN_QR(13) = "P"
        End If
    Else
        INSRREQUISITN_QR(11) = Null
        INSRREQUISITN_QR(12) = Null
        INSRREQUISITN_QR(13) = Null
End If
End If
If Trim(cmb_vpp.Text) = "NONE" Or Len(Trim(cmb_vpp.Text)) = 0 Then
    INSRREQUISITN_QR(14) = 0
End If
If Trim(cmb_vpp.Text) = "VPP" Then
    INSRREQUISITN_QR(14) = 1
End If
If Trim(cmb_vpp.Text) = "DER" Then
    INSRREQUISITN_QR(14) = 2
End If
If Trim(cmb_vpp.Text) = "BAR" Then
    INSRREQUISITN_QR(14) = 3
End If
If Trim(cmb_vpp.Text) = "DrOp" Then
    INSRREQUISITN_QR(14) = 4
End If
INSRREQUISITN_QR.Execute
    '*** INSERTING INT0 RREQUISITION
    '*** INSERTING INTO DIHREPORT
    If CMB_TEXT.ListCount > 0 Then
        For I = 0 To CMB_TEXT.ListCount - 1
            INSDIHREPORT_QR(0) = req_chglen(Trim(txt_ihreqno.Text))
            INSDIHREPORT_QR(1) = CMB_TEXT.List(I)
            INSDIHREPORT_QR.Execute
        Next
    End If
    '*** INSERTING INTO DIHREPORT
    '**** INSERTING INTO DREQUISITION

INSIHREQ_QR(0) = req_chglen(Trim(txt_ihreqno.Text))
INSIHREQ_QR(1) = Trim(txt_ihpathno.Text)
INSIHREQ_QR(2) = Trim(Msk_caseno.FormattedText)

INSIHREQ_QR(3) = Trim(cmb_Ihsoutsideblock.Text)
INSIHREQ_QR(4) = Trim(cmb_Ihsoutslide.Text)
INSIHREQ_QR(5) = "Y"
' UNCOMMENT AFTER ADDING THE FIELDS IN THE DATA BASE
'INSDREQUISITN_QR(15) = Trim(txt_opprocedure.Text)
INSIHREQ_QR(6) = Trim(cmb_refdoc.Text)
INSIHREQ_QR(7) = Trim(cmb_reftodoc.Text)
If txt_status = "IP" Then
    INSIHREQ_QR(8) = Trim(txt_ward.Text)
    INSIHREQ_QR(9) = Trim(txt_bedno.Text)
Else
    INSIHREQ_QR(8) = Null
    INSIHREQ_QR(9) = Null
End If
durill = txt_year & "/" & Trim(CMB_DURILLMONTH.Text)
INSIHREQ_QR(10) = durill
If IHFLAG = True Then
    INSIHREQ_QR(11) = req_chglen(Left(txt_sppathreqno.Text, 16))
Else
    INSIHREQ_QR(11) = ""
End If
If Len(Trim(txt_ihmeterial.Text)) > 0 Then
    INSIHREQ_QR(12) = Null
    INSIHREQ_QR(12) = Trim(txt_ihmeterial.Text)
Else
    INSIHREQ_QR(12) = Null
    INSIHREQ_QR(12) = " "
End If
INSIHREQ_QR.Execute
'**** INSERTING INTO DREQUISITION

 '*** UPDATING THE PROVISIONAL DIAGNOSIS
If SELTA1012FLAG = False Then
    selta1012_qr(0) = Msk_caseno.FormattedText
    Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
    SELTA1012FLAG = True
Else
     selta1012_qr(0) = Msk_caseno.FormattedText
     selta1012_rs.MoreResults
     selta1012_rs.Requery
End If
If selta1012_rs.RowCount > 0 Then
    UPDTA1012_QR(0) = Trim(txt_indctinvstg.Text)
    UPDTA1012_QR(1) = Msk_caseno.FormattedText
    UPDTA1012_QR.Execute
Else
    insta1012_qr(0) = Msk_caseno.FormattedText
    insta1012_qr(1) = Trim(txt_indctinvstg.Text)
    insta1012_qr.Execute
End If
'*** UPDATING THE PROVISIONAL DIAGNOSIS
'*** INSERTING INTO DMULPROCE FOR COLLECTION AND DEPT COLLECTION SCREEN
''INSDMULPROCE_QR(0) = req_chglen(txt_spreqno.Text)
''INSDMULPROCE_QR(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
''INSDMULPROCE_QR(2) = Msk_caseno.FormattedText
''INSDMULPROCE_QR(3) = cmb_procedurecode.Text
''INSDMULPROCE_QR.Execute

'***INSERTING INTO DMULPROCE FOR COLLECTION AND DEPT COLLECTION SCREEN
  
'*** INSERTING INTO DIHREPORT
    

'*** INSERTING INTO DIHREPORT
'If SELCLINHIST_RS.RowCount > 0 Then
'    updclinhist_qr(0) = Trim(txt_clinhist.Text)
'    updclinhist_qr(1) = Trim(Msk_caseno.FormattedText)
'    updclinhist_qr.Execute
'Else
    insclinhist_qr(0) = Trim(Msk_caseno.FormattedText)
    insclinhist_qr(1) = GET_REVDATE(Trim(DATE11))
    insclinhist_qr(2) = Trim(time1)
    If Trim(txt_clinhist.Text) = "NYD" Then
        insclinhist_qr(3) = "NYD"
    Else
        insclinhist_qr(3) = Trim(txt_clinhist.Text)
    End If
    insclinhist_qr.Execute
'End If

Data1.UpdateRecord
Data1.Recordset.Bookmark = Data1.Recordset.LastModified
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Public Sub NO_CHARGEIH()
On Error GoTo MSG:
No_Chargeflag = False
 
If SELMA101FLAG = False Then
    SELMA101_QR(0) = Trim(cmb_ihprocode.Text)
    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
    SELMA101FLAG = True
 Else
    SELMA101_QR(0) = Trim(cmb_ihprocode.Text)
    SELMA101_RS.MoreResults
    SELMA101_RS.Requery
End If
 
 
'SELMA101_QR(0) = Trim(cmb_ihprocode.Text)
'SELMA101_RS.MoreResults
'SELMA101_RS.Requery
If SELMA101_RS.RowCount > 0 Then
    If Trim(txt_catgory.Text) = "C" Then
        If SELMA101_RS("C_CH") = 0 Then
            No_Chargeflag = True
        End If
    ElseIf Trim(txt_catgory.Text) = "NC" Then
        If SELMA101_RS("NC_CH") = 0 Then
            No_Chargeflag = True
            NoMemo = True
        End If
    End If
Else
    'MsgBox "There is NO  RECORD with this SERVICE CODE."  'COMMENTED BY SHWETA 18-AUG-09
End If
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Public Sub IHENABL(str As String)
If str = "DIS" Then
    cmb_ihprocedure.Enabled = False
    cmb_Ihsoutsideblock.Enabled = False
    cmb_Ihspreinv.Enabled = False
    cmb_Ihsoutsideblock.Enabled = False
    cmb_ihprocedure.Enabled = False
    txt_sppathreqno.Enabled = False
    txt_ihmeterial.Enabled = False
End If
If str = "ENA" Then
    cmb_ihprocedure.Enabled = True
    cmb_Ihsoutsideblock.Enabled = True
    cmb_Ihspreinv.Enabled = True
    cmb_Ihsoutsideblock.Enabled = True
    cmb_ihprocedure.Enabled = True
    txt_sppathreqno.Enabled = True
    txt_ihmeterial.Enabled = True
End If
If str = "MOD" Then
    cmb_ihprocedure.Enabled = False
    cmb_Ihsoutsideblock.Enabled = True
    cmb_Ihspreinv.Enabled = True
    cmb_Ihsoutsideblock.Enabled = True
    cmb_ihprocedure.Enabled = False
    txt_sppathreqno.Enabled = False
    lbl_ihpathno.Visible = False
    lbl_ihreqno.Visible = False
    txt_ihpathno.Visible = False
    txt_ihreqno.Visible = False
    txt_indctinvstg.Enabled = True
    txt_clinhist.Enabled = True
    txt_ihmeterial.Enabled = True
End If
End Sub

Public Sub IHCLEAR()
    cmb_ihprocedure.Text = ""
    cmb_Ihsoutsideblock.clear
    cmb_Ihspreinv.clear
    cmb_Ihsoutslide.clear
    txt_sppathreqno.Text = ""
    txt_ihpathno.Text = " "
    cmb_Ihspreinv.clear
    cmb_ihprocode.clear
    cmb_ihprocedure.clear
    txt_ihmeterial.Text = ""
End Sub

Private Sub Opt_patient_Click()
If opt_patient.Value = True Then
    lbl_caseno.Caption = "Case No."
    Msk_caseno.Mask = "??>/#####"
    LBL_DEPENDENT.Visible = False
    cmb_dependents.Visible = False
     Msk_caseno.Enabled = True
     Msk_caseno.SetFocus
'    msk_caseno.SetFocus
   
End If
End Sub

Private Sub opt_staff_Click()
On Error GoTo MSG:
If opt_staff.Value = True Then

    lbl_caseno.Caption = "  CC No."
    Msk_caseno.Mask = "########"
    LBL_DEPENDENT.Visible = True
    cmb_dependents.Visible = True
    Msk_caseno.Enabled = True
    Msk_caseno.SetFocus
   
Else
    lbl_caseno.Caption = "Case No."
End If
'btnenabl ("NEW")
Exit Sub
MSG:
MsgBox Err.Description

End Sub
Private Sub Timer1_Timer()
    DATETIME
    txt_time.Text = Format(time1, "HH:MM:SS")
End Sub

Private Sub txt_cnatmeterial_GotFocus()
On Error GoTo m1:
WMETERIALCN.Visible = False
Set DMETERIALCN = WMETERIALCN.ActiveDocument
DMETERIALCN.Content = ""
DMETERIALCN.Content = txt_cnatmeterial
WMETERIALCN.Caption = "Nature of Meterial"
WMETERIALCN.WindowState = wdWindowStateMaximize
DMETERIALCN.Content = txt_cnatmeterial
WMETERIALCN.Visible = True
WMETERIALCN.Activate
m1:
On Error Resume Next
Set WMETERIALCN = CreateObject("Word.Application")
WMETERIALCN.Documents.add ("C:\DISREPORTSEC\REPORT.DOT")
WMETERIALCN.Caption = "Nature of Meterial"
WMETERIALCN.WindowState = wdWindowStateMaximize
Set DMETERIALCN = WMETERIALCN.ActiveDocument
DMETERIALCN.Content = Trim(txt_cnatmeterial.Text)
WMETERIALCN.Visible = True
WMETERIALCN.Activate
End Sub

Private Sub txt_cnatmeterial_LostFocus()
On Error GoTo m2:
Set DMETERIALCN = WMETERIALCN.ActiveDocument
txt_cnatmeterial.Text = DMETERIALCN.Content
If Len(Trim(txt_cnatmeterial.Text)) > 1000 Then
    MsgBox "Nature of Meterial Max of 1000 Characters"
    txt_cnatmeterial.SetFocus
    Exit Sub
End If
'flagimp = False
DMETERIALCN.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'DMETERIALCN.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
DMETERIALCN.Close (wdDoNotSaveChanges)
WMETERIALCN.quit
'W1.Visible = False
Exit Sub
m2:
MsgBox Err.Description
End Sub

Private Sub txt_fnnmeterial_GotFocus()
On Error GoTo m1:
WMETERIALFN.Visible = False
Set DMETERIALFN = WMETERIALFN.ActiveDocument
DMETERIALFN.Content = ""
DMETERIALFN.Content = txt_fnnmeterial
WMETERIALFN.Caption = "Nature of Meterial"
WMETERIALFN.WindowState = wdWindowStateMaximize
DMETERIALFN.Content = txt_fnnmeterial
WMETERIALFN.Visible = True
WMETERIALFN.Activate
m1:
On Error Resume Next
Set WMETERIALFN = CreateObject("Word.Application")
WMETERIALFN.Documents.add ("C:\DISREPORTSEC\REPORT.DOT")
WMETERIALFN.Caption = "Nature of Meterial"
WMETERIALFN.WindowState = wdWindowStateMaximize
Set DMETERIALFN = WMETERIALFN.ActiveDocument
DMETERIALFN.Content = Trim(txt_fnnmeterial.Text)
WMETERIALFN.Visible = True
WMETERIALFN.Activate
End Sub

Private Sub txt_fnnmeterial_LostFocus()
On Error GoTo m2:
Set DMETERIALFN = WMETERIALFN.ActiveDocument
txt_fnnmeterial.Text = DMETERIALFN.Content
If Len(Trim(txt_fnnmeterial.Text)) > 1000 Then
    MsgBox "Nature of Meterial Max of 1000 Characters"
    txt_fnnmeterial.SetFocus
    Exit Sub
End If
'flagimp = False
DMETERIALFN.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'DMETERIALFN.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
DMETERIALFN.Close (wdDoNotSaveChanges)
WMETERIALFN.quit
'W1.Visible = False
Exit Sub
m2:
MsgBox Err.Description
End Sub

Private Sub txt_ihmeterial_GotFocus()
On Error GoTo m1:
WMETERIALIH.Visible = False
Set DMETERIALIH = WMETERIALIH.ActiveDocument
DMETERIALIH.Content = ""
DMETERIALIH.Content = txt_ihmeterial
WMETERIALIH.Caption = "Nature of Meterial"
WMETERIALIH.WindowState = wdWindowStateMaximize
DMETERIALIH.Content = txt_ihmeterial
WMETERIALIH.Visible = True
WMETERIALIH.Activate
m1:
On Error Resume Next
Set WMETERIALIH = CreateObject("Word.Application")
WMETERIALIH.Documents.add ("C:\DISREPORTSEC\REPORT.DOT")
WMETERIALIH.Caption = "Nature of Meterial"
WMETERIALIH.WindowState = wdWindowStateMaximize
Set DMETERIALIH = WMETERIALIH.ActiveDocument
DMETERIALIH.Content = Trim(txt_ihmeterial.Text)
WMETERIALIH.Visible = True
WMETERIALIH.Activate
End Sub

Private Sub txt_ihmeterial_LostFocus()
On Error GoTo m2:
Set DMETERIALIH = WMETERIALIH.ActiveDocument
txt_ihmeterial.Text = DMETERIALIH.Content
If Len(Trim(txt_ihmeterial.Text)) >= 1000 Then
    MsgBox "Nature of Meterial Max. of 1000 Characters"
    txt_ihmeterial.SetFocus
    Exit Sub
End If
'flagimp = False
DMETERIALIH.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'DMETERIALIH.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
DMETERIALIH.Close (wdDoNotSaveChanges)
WMETERIALIH.quit
'W1.Visible = False
Exit Sub
m2:
MsgBox Err.Description
End Sub


'*************** Commented by Shraddha Pandit on 29/04/2009 ************
'Private Sub txt_mbspecimen_LostFocus()
'   txt_mbspecimen.Text = UCase(txt_mbspecimen.Text)
'End Sub
'***********************************************************************

Private Sub txt_nmeterial_DblClick()
''    On Error GoTo m1:
''NM_FLG = False
''WMETERIAL.Visible = False
''Set DMETERIAL = WMETERIAL.ActiveDocument
''DMETERIAL.Content = ""
''DMETERIAL.Content = txt_nmeterial
''WMETERIAL.Caption = "Nature of Meterial"
''WMETERIAL.WindowState = wdWindowStateMaximize
''DMETERIAL.Content = txt_nmeterial
''WMETERIAL.Visible = True
''WMETERIAL.Activate
''m1:
''On Error Resume Next
''NM_FLG = False
''Set WMETERIAL = CreateObject("Word.Application")
''WMETERIAL.Documents.add ("C:\DISREPORTSEC\REPORT.DOT")
''WMETERIAL.Caption = "Nature of Meterial"
''WMETERIAL.WindowState = wdWindowStateMaximize
''Set DMETERIAL = WMETERIAL.ActiveDocument
''DMETERIAL.Content = Trim(txt_nmeterial.Text)
''WMETERIAL.Visible = True
''WMETERIAL.Activate
End Sub

Private Sub txt_nmeterial_GotFocus()
 On Error GoTo m1:
WMETERIAL.Visible = False
Set DMETERIAL = WMETERIAL.ActiveDocument
DMETERIAL.Content = ""
DMETERIAL.Content = txt_nmeterial
WMETERIAL.Caption = "Nature of Meterial"
WMETERIAL.WindowState = wdWindowStateMaximize
DMETERIAL.Content = txt_nmeterial
WMETERIAL.Visible = True
WMETERIAL.Activate
m1:
On Error Resume Next
Set WMETERIAL = CreateObject("Word.Application")
WMETERIAL.Documents.add ("C:\DISREPORTSEC\REPORT.DOT")
WMETERIAL.Caption = "Nature of Meterial"
WMETERIAL.WindowState = wdWindowStateMaximize
Set DMETERIAL = WMETERIAL.ActiveDocument
DMETERIAL.Content = Trim(txt_nmeterial.Text)
WMETERIAL.Visible = True
WMETERIAL.Activate
'NM_FLG = True
End Sub

Private Sub txt_nmeterial_LostFocus()
On Error GoTo m2:
  '  If NM_FLG = True Then
   '
   '     NM_FLG = False
   '     Exit Sub
   ' End If

Set DMETERIAL = WMETERIAL.ActiveDocument
txt_nmeterial.Text = DMETERIAL.Content
If Len(Trim(txt_nmeterial.Text)) >= 1000 Then
    MsgBox "Nature of Meterial Max. of 1000 Characters"
    txt_nmeterial.SetFocus
    Exit Sub
End If
'flagimp = False
DMETERIAL.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'DMETERIAL.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
DMETERIAL.Close (wdDoNotSaveChanges)
WMETERIAL.quit
'W1.Visible = False
Exit Sub
m2:
MsgBox Err.Description
End Sub

Public Sub SPUPD()
'**** Updating the drequisition table
        upddrequisitn_qr(0) = (cmb_refdoc.Text)
        upddrequisitn_qr(1) = cmb_reftodoc.Text
        upddrequisitn_qr(2) = txt_year & "/" & CMB_DURILLMONTH
        I = InStr(cmb_sites, "$")
        If I = 0 Then
               If sitedupflag = False Then
                    sitedupqr(0) = UCase(Trim(cmb_sites.Text))
                    Set siteduprs = sitedupqr.OpenResultset(1, 3)
                    sitedupflag = True
               Else
                    sitedupqr(0) = UCase(Trim(cmb_sites.Text))
                    siteduprs.MoreResults
                    siteduprs.Requery
                                   
               End If
               If siteduprs.RowCount > 0 Then
                    upddrequisitn_qr(3) = siteduprs("SITECD")
                          
               Else
               
                    If sitedup1flag = False Then
                        'sitedupqr(0) = Trim(cmb_sites.Text)
                        Set siteduprs1 = sitedupqr1.OpenResultset(1, 3)
                        sitedup1flag = True
                    Else
                        siteduprs1.MoreResults
                        siteduprs1.Requery
                    End If
                    If (siteduprs1.RowCount > 0) Then
                        
                        Dim a As Integer
                        If Len(Trim(cmb_sites.Text)) > 0 Then
                            a = siteduprs1(0)
                            a = a + 1
                            siteinsertqr(0) = a
                            siteinsertqr(1) = UCase(Trim(cmb_sites.Text))
                            siteinsertqr.Execute
                            upddrequisitn_qr(3) = a
                         Else
                            upddrequisitn_qr(3) = "9999"
                         End If
                        
                     Else
                        
                        a = 5000
                        siteinsertqr(0) = a
                        siteinsertqr(1) = UCase(Trim(cmb_sites.Text))
                        siteinsertqr.Execute
                        upddrequisitn_qr(3) = a
                     End If
                End If
         Else
               upddrequisitn_qr(3) = Mid(cmb_sites, I + 1, 4)
         End If
        
        
        
        upddrequisitn_qr(4) = UCase(Trim(cmb_laterality))
        upddrequisitn_qr(5) = UCase(Trim(TXT_OPPROCEDURE.Text))
        If Len(Trim(txt_nmeterial.Text)) > 0 Then
            upddrequisitn_qr(6) = Null
            upddrequisitn_qr(6) = UCase(Trim(txt_nmeterial.Text))
        Else
            upddrequisitn_qr(6) = Null
            upddrequisitn_qr(6) = " "
        End If
        upddrequisitn_qr(7) = UCase(Trim(txt_metastasis.Text))
        upddrequisitn_qr(8) = UCase(Trim(cmb_laterality1.Text))
        upddrequisitn_qr(9) = cmb_outslide
        upddrequisitn_qr(10) = txt_outspecimen.Text
        upddrequisitn_qr(11) = cmb_outblock
        If Trim(cmb_frozensec.Text) = "NO" Or Trim(cmb_frozensec.Text) = "" Then
            upddrequisitn_qr(12) = "NO"
            upddrequisitn_qr(13) = Null
        Else
            upddrequisitn_qr(12) = UCase(Trim(cmb_frozensec.Text))
            upddrequisitn_qr(13) = GET_REVDATE(MSK_FROSECDT.FormattedText)
        End If
        If Len(Trim(cmb_spfnac.Text)) = 0 Then
            upddrequisitn_qr(14) = Null
        Else
            upddrequisitn_qr(14) = cmb_spfnac.Text
        End If
        upddrequisitn_qr(15) = req_chglen(cmb_reqno.Text)
        upddrequisitn_qr.Execute
        '**** Updating the drequisition table
         '*** UPDATING THE PROVISIONAL DIAGNOSIS
        If opt_patient.Value = True Then
        If SELTA1012FLAG = False Then
            selta1012_qr(0) = Msk_caseno.FormattedText
            Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
            SELTA1012FLAG = True
        Else
             selta1012_qr(0) = Msk_caseno.FormattedText
             selta1012_rs.MoreResults
             selta1012_rs.Requery
        End If
        If selta1012_rs.RowCount > 0 Then
            UPDTA1012_QR(0) = Trim(txt_indctinvstg.Text)
            UPDTA1012_QR(1) = Msk_caseno.FormattedText
            UPDTA1012_QR.Execute
        Else
            insta1012_qr(0) = Msk_caseno.FormattedText
            insta1012_qr(1) = Trim(txt_indctinvstg.Text)
            insta1012_qr.Execute
        End If
        End If
        '*** UPDATING THE PROVISIONAL DIAGNOSIS
        '*** updating the clinical history
        If SELCLINHISTFLAG = False Then
            SELCLINHIST_QR(0) = Msk_caseno.FormattedText
            Set SELCLINHIST_RS = SELCLINHIST_QR.OpenResultset(1, 3)
            SELCLINHISTFLAG = True
        Else
            SELCLINHIST_QR(0) = Msk_caseno.FormattedText
            SELCLINHIST_RS.MoreResults
            SELCLINHIST_RS.Requery
        End If
        'If SELCLINHIST_RS.RowCount > 0 Then
'            updclinhist_qr(0) = Trim(txt_clinhist.Text)
'            updclinhist_qr(1) = Trim(Msk_caseno.FormattedText)
'            updclinhist_qr.Execute
'        Else
            insclinhist_qr(0) = Trim(Msk_caseno.FormattedText)
            insclinhist_qr(1) = GET_REVDATE(Trim(DATE11))
            insclinhist_qr(2) = Trim(time1)
            If Trim(txt_clinhist.Text) = "NYD" Then
                insclinhist_qr(3) = "NYD"
            Else
                insclinhist_qr(3) = Trim(txt_clinhist.Text)
            End If
            insclinhist_qr.Execute
        'End If
        '*** UPDATING THE RREQUISITN TABLE
        updrrequisitn_qr(0) = UCase(Trim(cmb_refdoc.Text))
        updrrequisitn_qr(1) = req_chglen(cmb_reqno.Text)
        updrrequisitn_qr.Execute
        
         If SELRREQUISITNFLAG2 = False Then
                SELRREQUISITN_QR2(0) = req_chglen(Trim(cmb_reqno.Text))
                Set SELRREQUISITN_RS2 = SELRREQUISITN_QR2.OpenResultset(1, 3)
                SELRREQUISITNFLAG2 = True
           Else
                SELRREQUISITN_QR2(0) = req_chglen(Trim(cmb_reqno.Text))
                SELRREQUISITN_RS2.MoreResults
                SELRREQUISITN_RS2.Requery
           End If
        If SELRREQUISITN_RS2.RowCount > 0 Then
                If IsNull(SELRREQUISITN_RS2("reqnstat")) Then
        
                If cmb_vpp.Text = "VPP" Or cmb_vpp.Text = "DER" Or cmb_vpp.Text = "BAR" Or cmb_vpp.Text = "DrOp" Or cmb_vpp.Text = "CHIP" Then
                    If cmb_vpp.Text = "VPP" Then
                        UPDRREQQR(0) = Null
                        UPDRREQQR(1) = Null
                        UPDRREQQR(2) = Null
                        UPDRREQQR(3) = 1
                        UPDRREQQR(4) = req_chglen(cmb_reqno.Text)
                        UPDRREQQR.Execute
                    End If
                    If cmb_vpp.Text = "DER" Then
                        UPDRREQQR(0) = "DER"
                        UPDRREQQR(1) = GET_REVDATE(Trim(DATE11))
                        UPDRREQQR(2) = "P"
                        UPDRREQQR(3) = 2
                        UPDRREQQR(4) = req_chglen(cmb_reqno.Text)
                        UPDRREQQR.Execute
                    End If
                    If cmb_vpp.Text = "BAR" Then
                        UPDRREQQR(0) = "BAR"
                        UPDRREQQR(1) = GET_REVDATE(Trim(DATE11))
                        UPDRREQQR(2) = "P"
                        UPDRREQQR(3) = 3
                        UPDRREQQR(4) = req_chglen(cmb_reqno.Text)
                        UPDRREQQR.Execute
                    End If
                    If cmb_vpp.Text = "DrOp" Then
                        UPDRREQQR(0) = "DrOp"
                        UPDRREQQR(1) = GET_REVDATE(Trim(DATE11))
                        UPDRREQQR(2) = "P"
                        UPDRREQQR(3) = 4
                        UPDRREQQR(4) = req_chglen(cmb_reqno.Text)
                        UPDRREQQR.Execute
                    End If
''''                    If cmb_vpp.Text = "CHIP" Then
''''                        UPDRREQQR(0) = "CHIP"
''''                        UPDRREQQR(1) = GET_REVDATE(Trim(DATE11))
''''                        UPDRREQQR(2) = "P"
''''                        UPDRREQQR(3) = 5
''''                        UPDRREQQR(4) = req_chglen(cmb_reqno.Text)
''''                        UPDRREQQR.Execute
''''                    End If
                
                End If
              End If
       End If
        '*** UPDATING THE RREQUISITN TABLE
End Sub

Public Sub IHUPD()
    '*** UPDATING DREQUISITION TABLE
        UPDDREQIHQR(0) = cmb_Ihsoutslide.Text
        UPDDREQIHQR(1) = cmb_Ihsoutsideblock.Text
        If Len(Trim(txt_ihmeterial.Text)) > 0 Then
            UPDDREQIHQR(2) = Null
            UPDDREQIHQR(2) = Trim(txt_ihmeterial.Text)
        Else
            UPDDREQIHQR(2) = Null
            UPDDREQIHQR(2) = " "
        End If
        UPDDREQIHQR(3) = req_chglen(cmb_reqno.Text)
        UPDDREQIHQR.Execute

    '*** UPDATING DREQUISITION TABLE
 '*** UPDATING THE PROVISIONAL DIAGNOSIS
        If opt_patient.Value = True Then
        If SELTA1012FLAG = False Then
            selta1012_qr(0) = Msk_caseno.FormattedText
            Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
            SELTA1012FLAG = True
        Else
             selta1012_qr(0) = Msk_caseno.FormattedText
             selta1012_rs.MoreResults
             selta1012_rs.Requery
        End If
        If selta1012_rs.RowCount > 0 Then
            UPDTA1012_QR(0) = Trim(txt_indctinvstg.Text)
            UPDTA1012_QR(1) = Msk_caseno.FormattedText
            UPDTA1012_QR.Execute
        Else
            insta1012_qr(0) = Msk_caseno.FormattedText
            insta1012_qr(1) = Trim(txt_indctinvstg.Text)
            insta1012_qr.Execute
        End If
        End If
        '*** UPDATING THE PROVISIONAL DIAGNOSIS
        '*** updating the clinical history
        If SELCLINHISTFLAG = False Then
            SELCLINHIST_QR(0) = Msk_caseno.FormattedText
            Set SELCLINHIST_RS = SELCLINHIST_QR.OpenResultset(1, 3)
            SELCLINHISTFLAG = True
        Else
            SELCLINHIST_QR(0) = Msk_caseno.FormattedText
            SELCLINHIST_RS.MoreResults
            SELCLINHIST_RS.Requery
        End If
'        If SELCLINHIST_RS.RowCount > 0 Then
'            updclinhist_qr(0) = Trim(txt_clinhist.Text)
'            updclinhist_qr(1) = Trim(Msk_caseno.Text)
'            updclinhist_qr.Execute
'        Else
            insclinhist_qr(0) = Trim(Msk_caseno.FormattedText)
            insclinhist_qr(1) = GET_REVDATE(Trim(DATE11))
            insclinhist_qr(2) = Trim(time1)
            If Trim(txt_clinhist.Text) = "NYD" Then
                insclinhist_qr(3) = "NYD"
            Else
                insclinhist_qr(3) = Trim(txt_clinhist.Text)
            End If
            insclinhist_qr.Execute
       ' End If
        '*** updating the clinical history
        updrrequisitn_qr(0) = cmb_refdoc.Text
        updrrequisitn_qr(1) = req_chglen(cmb_reqno.Text)
        updrrequisitn_qr.Execute
         If SELRREQUISITNFLAG2 = False Then
                SELRREQUISITN_QR2(0) = req_chglen(Trim(cmb_reqno.Text))
                Set SELRREQUISITN_RS2 = SELRREQUISITN_QR2.OpenResultset(1, 3)
                SELRREQUISITNFLAG2 = True
           Else
                SELRREQUISITN_QR2(0) = req_chglen(Trim(cmb_reqno.Text))
                SELRREQUISITN_RS2.MoreResults
                SELRREQUISITN_RS2.Requery
           End If
        If SELRREQUISITN_RS2.RowCount > 0 Then
                If IsNull(SELRREQUISITN_RS2("reqnstat")) Then
        
                If CMB_IHNIP.Text = "VPP" Or CMB_IHNIP.Text = "LR" Or CMB_IHNIP.Text = "BAR" Or CMB_IHNIP.Text = "DrOp" Then
                    If CMB_IHNIP.Text = "VPP" Then
                        UPDRREQQR(0) = Null
                        UPDRREQQR(1) = Null
                        UPDRREQQR(2) = Null
                        UPDRREQQR(3) = 1
                        UPDRREQQR(4) = req_chglen(cmb_reqno.Text)
                        UPDRREQQR.Execute
                    End If
                    If CMB_IHNIP.Text = "LR" Then
                        UPDRREQQR(0) = "LR"
                        UPDRREQQR(1) = GET_REVDATE(Trim(DATE11))
                        UPDRREQQR(2) = "P"
                        UPDRREQQR(3) = 2
                        UPDRREQQR(4) = req_chglen(cmb_reqno.Text)
                        UPDRREQQR.Execute
                    End If
                    If CMB_IHNIP.Text = "BAR" Then
                        UPDRREQQR(0) = "BAR"
                        UPDRREQQR(1) = GET_REVDATE(Trim(DATE11))
                        UPDRREQQR(2) = "P"
                        UPDRREQQR(3) = 3
                        UPDRREQQR(4) = req_chglen(cmb_reqno.Text)
                        UPDRREQQR.Execute
                    End If
                    If CMB_IHNIP.Text = "DrOp" Then
                        UPDRREQQR(0) = "DrOp"
                        UPDRREQQR(1) = GET_REVDATE(Trim(DATE11))
                        UPDRREQQR(2) = "P"
                        UPDRREQQR(3) = 4
                        UPDRREQQR(4) = req_chglen(cmb_reqno.Text)
                        UPDRREQQR.Execute
                    End If
                End If
              End If
       End If
        
        
End Sub

Public Sub FA()
        surgicalPathology.Visible = False
        immunohistochemistry.Visible = False
        frm_hematopathology.Visible = False
        biochemistry.Visible = False
        bceltm.Visible = False
'        cardiology.Visible = False
        Microbiology.Visible = False
        fnac.Visible = True
        CYTOLOGY.Visible = False
        'cpl.Visible = False
        TUMOURMARKER.Visible = False
        hemcbc.Hide
        HEMBF.Hide
        HEMBMA.Hide
        'HLATYPING.Visible = False
        lbl_req.Left = 4000
        lbl_req.Caption = "REQUISITION - FNAC"
        pathology.Caption = "Pathology - Fine Needle Aspiration Cytology"
        FACLEAR
        FAENABL ("ENA")
        
        ' PROCEDURE CODE FILLING IN THE COMBOBOX
'        If SELRINVESTFLAG = False Then
'            SELRINVEST_QR(0) = Trim(txt_modcode.Text)
'            Set SELRINVEST_RS = SELRINVEST_QR.OpenResultset(1, 3)
'            SELRINVESTFLAG = True
'        Else
'            SELRINVEST_QR(0) = Trim(txt_modcode.Text)
'            SELRINVEST_RS.MoreResults
'            SELRINVEST_RS.Requery
'        End If
'        If SELRINVEST_RS.RowCount > 0 Then
'            cmb_procedure.clear
'            cmb_procedurecode.clear
'            While Not SELRINVEST_RS.EOF
'                cmb_fnprocode.AddItem Trim(SELRINVEST_RS("ser_code"))
'                cmb_fnprocode.Text = Trim(SELRINVEST_RS("ser_code"))
'                If SELMA101FLAG = False Then
'                    SELMA101_QR(0) = Trim(SELRINVEST_RS("SER_CODE"))
'                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
'                    SELMA101FLAG = True
'                 Else
'                    SELMA101_QR(0) = SELRINVEST_RS("SER_CODE")
'                    SELMA101_RS.MoreResults
'                    SELMA101_RS.Requery
'                End If
'                If SELMA101_RS.RowCount > 0 Then
'                    While Not SELMA101_RS.EOF
'                    CMB_FNPROCEDURE.AddItem SELMA101_RS("SER_DET")
'                    CMB_FNPROCEDURE.Text = SELMA101_RS("SER_DET")
'                    SELMA101_RS.MoveNext
'                    Wend
'                End If
'                SELRINVEST_RS.MoveNext
'            Wend
'        End If
        
        ''ADDED BY MONO ON 02/06/2003
    
    If SELRINVEST11_FLAG = False Then
        SELRINVEST11_QR(0) = Trim(txt_modcode.Text)
        Set SELRINVEST11_RS = SELRINVEST11_QR.OpenResultset(1, 3)
        SELRINVEST11_FLAG = True
    Else
        SELRINVEST11_QR(0) = Trim(txt_modcode.Text)
        SELRINVEST11_RS.MoreResults
        SELRINVEST11_RS.Requery
    End If
    If SELRINVEST11_RS.RowCount > 0 Then
        CMB_FNPROCEDURE.clear
        cmb_fnprocode.clear
        While Not SELRINVEST11_RS.EOF
            cmb_fnprocode.AddItem Trim(SELRINVEST11_RS("ser_code"))
            cmb_fnprocode.Text = Trim(SELRINVEST11_RS("ser_code"))
            'cmb_FNprocedurecode.AddItem Trim(SELRINVEST11_RS("ser_code"))
            If SELMA101_FLAG11 = False Then
                SELMA101_QR11(0) = Trim(SELRINVEST11_RS("SER_CODE"))
                Set SELMA101_RS11 = SELMA101_QR11.OpenResultset(1, 3)
                SELMA101_FLAG11 = True
             Else
                SELMA101_QR11(0) = SELRINVEST11_RS("SER_CODE")
                SELMA101_RS11.MoreResults
                SELMA101_RS11.Requery
            End If
            If SELMA101_RS11.RowCount > 0 Then
                While Not SELMA101_RS11.EOF
                CMB_FNPROCEDURE.AddItem SELMA101_RS11("SER_DET")
                CMB_FNPROCEDURE.Text = SELMA101_RS11("SER_DET")
                SELMA101_RS11.MoveNext
                Wend
            End If
            SELRINVEST11_RS.MoveNext
        Wend
    End If

        For I = 0 To 20
            TXT_FNSLIDEDRY.AddItem I
            TXT_FNOSLIDEWET.AddItem I
            cmb_outsideslidesfn.AddItem I
        Next I
    
    ' PROCEDURE CODE FILLING IN THE COMBOBOX
    ' SITES FILLING IN THE COMBOBOX
    If SELMA301FLAG = False Then
        Set SELMA301_RS = SELMA301_QR.OpenResultset(1, 3)
        SELMA301FLAG = True
    Else
        Set SELMA301_RS = SELMA301_QR.OpenResultset(1, 3)
        SELMA301_RS.MoreResults
        SELMA301_RS.Requery
    End If
    
    If SELMA301_RS.RowCount > 0 Then
        cmb_sites.clear
        While Not SELMA301_RS.EOF
        If Trim(SELMA301_RS("sitedesc")) <> "" Then
            cmb_fnsite.AddItem Trim(SELMA301_RS("SITEDESC")) & "$" & Trim(SELMA301_RS("SITECD"))
        End If
        SELMA301_RS.MoveNext
        Wend
    End If
    ' SITES FILLING IN THE COMBOBOX
    ' LATERALITY FILLING IN THE COMBOBOX
    If SELMA10122FLAG = False Then
        Set SELMA10122_RS = SELMA10122_QR.OpenResultset(1, 3)
        SELMA10122FLAG = True
    Else
    
        SELMA10122_RS.MoreResults
        SELMA10122_RS.Requery
    End If
    If SELMA10122_RS.RowCount > 0 Then
        cmb_fnlaterality1.clear
        cmb_fnlaterality.clear
        While Not SELMA10122_RS.EOF
        cmb_fnlaterality1.AddItem Trim(SELMA10122_RS("ADJ_PRVDIA"))
        cmb_fnlaterality.AddItem Trim(SELMA10122_RS("ADJ_PRVDIA"))
        SELMA10122_RS.MoveNext
        Wend
    End If
    ' LATERALITY FILLING IN THE COMBOBOX
    ' FILLING THE PREVIOUSLY INVESTIGATED REQ NO.
                If seLdreqrsflag = False Then
                    SELDREQQR(0) = Msk_caseno.FormattedText
                    Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                    seLdreqrsflag = True
                Else
                    SELDREQQR(0) = Msk_caseno.FormattedText
                    SELDREQRS.MoreResults
                    SELDREQRS.Requery
                End If
                If SELDREQRS.RowCount > 0 Then
                    cmb_preinvest.clear
                    While Not SELDREQRS.EOF
                        cmb_fnpreinvest.AddItem SELDREQRS("PATHNO") & "$" & SELDREQRS("REQNO")
                        SELDREQRS.MoveNext
                    Wend
                    cmb_fnpreinvest.ListIndex = 0
                End If

  ' FILLING THE PREVIOUSLY INVESTIGATED REQ NO.
'  CMB_CELLBLOCK.AddItem "YES"
'  CMB_CELLBLOCK.AddItem "NO"
'    If SELDENTERFLAG = False Then
'        Set SELDENTERRS = SELDENTERQR.OpenResultset(1, 3)
'        SELDENTERFLAG = True
'    Else
'        SELDENTERRS.MoreResults
'        SELDENTERRS.Requery
'    End If
'    If SELDENTERRS.RowCount > 0 Then
'        cmb_fnac.clear
'        While Not SELDENTERRS.EOF
'            cmb_fnac.AddItem Trim(SELDENTERRS("CCNAME"))
'            SELDENTERRS.MoveNext
'        Wend
'    End If
    If Left(Msk_caseno.FormattedText, 2) = "RF" Or Left(Msk_caseno.FormattedText, 2) = "RG" Then
        cmb_outsideslidesfn.Visible = True
        lbl_outsideslides.Visible = True
    Else
        cmb_outsideslidesfn.Visible = False
        lbl_outsideslides.Visible = False
    End If
    lbl_fnreqno.Visible = True
    txt_FNreqno.Visible = True
    
    If SELMA1031FLAG = False Then
            Set SELMA1031_RS = SELMA1031_QR.OpenResultset(1, 1)
            SELMA1031FLAG = True
        Else
            Set SELMA1031_RS = SELMA1031_QR.OpenResultset(1, 1)
            SELMA1031_RS.MoreResults
            SELMA1031_RS.Requery
    End If
        If SELMA1031_RS.RowCount > 0 Then
            cmb_fnac.clear
            While Not SELMA1031_RS.EOF
             'No need of performed by here so commented
             'by Shraddha Pandit 03/11/08
                'cmb_fnac.AddItem Trim(SELMA1031_RS("EMPTITLE"))
                'cmb_fnac.AddItem Trim(SELMA1031_RS("UD_NAME"))
                
                SELMA1031_RS.MoveNext
            Wend
            
        End If
End Sub

Public Sub cg()
        surgicalPathology.Visible = False
        immunohistochemistry.Visible = False
        frm_hematopathology.Visible = False
        biochemistry.Visible = False
        bceltm.Visible = False
'        cardiology.Visible = False
        Microbiology.Visible = False
        fnac.Visible = False
        CYTOLOGY.Visible = True
        hemcbc.Hide
        HEMBF.Hide
        HEMBMA.Hide
        If txt_modcode.Text = "CG" Then
            CYTOLOGY.Caption = "Gynaec"
            
            lbl_req.Caption = "REQUISITION - CYTOLOGY GYNAEC"
            pathology.Caption = "Pathology - Cytology Gynaec"
            lbl_lmp.Visible = True
            MSK_LMP.Visible = True
            lbl_scopyfind.Visible = True
            txt_scopyfind.Visible = True
            If Trim(txt_sex.Text) = "F" Then
                lbl_lmp.Visible = True
                MSK_LMP.Visible = True
            Else
                lbl_lmp.Visible = False
                MSK_LMP.Visible = False
            End If
       End If
       If txt_modcode.Text = "CN" Then
            CYTOLOGY.Caption = "Non - Gynaec"
            
            lbl_req.Caption = "REQUISITION - CYTOLOGY NONGYNAEC"
            pathology.Caption = "Pathology - Cytology NonGynaec"
            lbl_lmp.Visible = False
            MSK_LMP.Visible = False
            lbl_scopyfind.Visible = True
            txt_scopyfind.Visible = True
            lbl_lmp.Visible = False
            MSK_LMP.Visible = False
        End If
        
        
        If txt_modcode.Text = "PG" Then
            CYTOLOGY.Caption = "Preventive Gynaec"
            
            lbl_req.Caption = "REQUISITION - PREVENTIVE GYNAEC"
            pathology.Caption = "Pathology - Preventive Gynaec"
            lbl_lmp.Visible = False
            MSK_LMP.Visible = False
            lbl_scopyfind.Visible = True
            txt_scopyfind.Visible = True
            lbl_lmp.Visible = False
            MSK_LMP.Visible = False
        End If
        
        If txt_modcode.Text = "PN" Then
            CYTOLOGY.Caption = "Preventive Non-Gynaec"
            
            lbl_req.Caption = "REQUISITION - PREVENTIVE NON GYNAEC"
            pathology.Caption = "Pathology - Preventive Gynaec"
            lbl_lmp.Visible = False
            MSK_LMP.Visible = False
            lbl_scopyfind.Visible = True
            txt_scopyfind.Visible = True
            lbl_lmp.Visible = False
            MSK_LMP.Visible = False
        End If
        
        
        
        'cpl.Visible = False
        TUMOURMARKER.Visible = False
        'HLATYPING.Visible = False
        lbl_req.Left = 3500
        CGENABL ("ENA")
        CGCLEAR
        ' PROCEDURE CODE FILLING IN THE COMBOBOX
    If SELRINVESTFLAG = False Then
        SELRINVEST_QR(0) = Trim(txt_modcode.Text)
        Set SELRINVEST_RS = SELRINVEST_QR.OpenResultset(1, 3)
        SELRINVESTFLAG = True
    Else
        SELRINVEST_QR(0) = Trim(txt_modcode.Text)
        SELRINVEST_RS.MoreResults
        SELRINVEST_RS.Requery
    End If
    If SELRINVEST_RS.RowCount > 0 Then
        CMB_CGPROCEDURE.clear
        cbo_proccode.clear
        While Not SELRINVEST_RS.EOF
            cbo_proccode.AddItem Trim(SELRINVEST_RS("ser_code"))
            'cbo_proccode.Text = SELRINVEST_RS("ser_code")
            If SELMA101FLAG = False Then
                SELMA101_QR(0) = Trim(SELRINVEST_RS("SER_CODE"))
                Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                SELMA101FLAG = True
             Else
                SELMA101_QR(0) = SELRINVEST_RS("SER_CODE")
                SELMA101_RS.MoreResults
                SELMA101_RS.Requery
            End If
            If SELMA101_RS.RowCount > 0 Then
                While Not SELMA101_RS.EOF
                CMB_CGPROCEDURE.AddItem SELMA101_RS("SER_DET")
                'CMB_CGPROCEDURE.Text = SELMA101_RS("SER_DET")
                SELMA101_RS.MoveNext
                Wend
            End If
            SELRINVEST_RS.MoveNext
        Wend
        
        
    End If
    
    ' PROCEDURE CODE FILLING IN THE COMBOBOX
    ' SITES FILLING IN THE COMBOBOX
    If txt_modcode.Text = "CN" Then
        If SELMA301FLAG = False Then
            Set SELMA301_RS = SELMA301_QR.OpenResultset(1, 3)
            SELMA301FLAG = True
        Else
            Set SELMA301_RS = SELMA301_QR.OpenResultset(1, 3)
            SELMA301_RS.MoreResults
            SELMA301_RS.Requery
        End If
        
        If SELMA301_RS.RowCount > 0 Then
            cmb_cnsite.clear
            While Not SELMA301_RS.EOF
            If Trim(SELMA301_RS("sitedesc")) <> "" Then
                cmb_cnsite.AddItem Trim(SELMA301_RS("SITEDESC")) & "$" & Trim(SELMA301_RS("SITECD"))
            End If
            SELMA301_RS.MoveNext
            Wend
        End If
    End If
    If txt_modcode.Text = "CG" Then
        cmb_cnsite.clear
        cmb_cnsite.AddItem "VULVA"
        cmb_cnsite.AddItem "VAGINA"
        cmb_cnsite.AddItem "CERVIX"
        cmb_cnsite.AddItem "ENDOMETRIUM"
        cmb_cnsite.AddItem "PERITIONEAL FLUID"
        cmb_cnsite.AddItem "INGUINAL NODE"
        cmb_cnsite.AddItem "OTHERS"
    End If
    ' SITES FILLING IN THE COMBOBOX
    ' LATERALITY FILLING IN THE COMBOBOX
    If SELMA10122FLAG = False Then
        Set SELMA10122_RS = SELMA10122_QR.OpenResultset(1, 3)
        SELMA10122FLAG = True
    Else
    
        SELMA10122_RS.MoreResults
        SELMA10122_RS.Requery
    End If
    If SELMA10122_RS.RowCount > 0 Then
        cmb_claterality.clear
        
        While Not SELMA10122_RS.EOF
             cmb_claterality.AddItem Trim(SELMA10122_RS("ADJ_PRVDIA"))
        SELMA10122_RS.MoveNext
        Wend
    End If
    ' LATERALITY FILLING IN THE COMBOBOX
  
  
  ' FILLING THE PREVIOUSLY INVESTIGATED REQ NO.
   
'   If SELRREQUISITNFLAG = False Then
'        SELRREQUISITN_QR(0) = Msk_caseno.FormattedText
'        Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
'        SELRREQUISITNFLAG = True
'   Else
'
'
'        SELRREQUISITN_QR(0) = Msk_caseno.FormattedText
'        SELRREQUISITN_RS.MoreResults
'        SELRREQUISITN_RS.Requery
'
'    End If
'    If SELRREQUISITN_RS.RowCount > 0 Then
'        cmb_preinvest.clear
'        While Not SELRREQUISITN_RS.EOF
'            cmb_preinvest.AddItem SELRREQUISITN_RS("REQNNO")
'            SELRREQUISITN_RS.MoveNext
'        Wend
'        cmb_preinvest.ListIndex = 0
'    End If
                    'ADDED BY MILIND ON 01112006 FOR DISPLAYING OT PROCEDUR NO.
                    If SELSCOPYFNDFLAG = False Then
                        SELSCOPYFNDQR(0) = Msk_caseno.FormattedText
                        Set SELSCOPYFNDRS = SELSCOPYFNDQR.OpenResultset(1, 3)
                        SELSCOPYFNDFLAG = True
                    Else
                        SELSCOPYFNDQR(0) = Msk_caseno.FormattedText
                        SELSCOPYFNDRS.MoreResults
                        SELSCOPYFNDRS.Requery
                    End If
                    
                    If SELSCOPYFNDRS.RowCount > 0 Then
                        If SELSCOPYFNDRS.EOF = False Then
                            txt_scopyfind.Text = SELSCOPYFNDRS("SURGPROCNO")
                        Else
                            txt_scopyfind.Text = ""
                        End If
                    End If
                    'UPTO HERE
                
                If seLdreqrsflag = False Then
                    SELDREQQR(0) = Msk_caseno.FormattedText
                    Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                    seLdreqrsflag = True
                Else
                    SELDREQQR(0) = Msk_caseno.FormattedText
                    SELDREQRS.MoreResults
                    SELDREQRS.Requery
                End If
                If SELDREQRS.RowCount > 0 Then
                    cmb_preinvest.clear
                    While Not SELDREQRS.EOF
                        cmb_cpreinv.AddItem SELDREQRS("PATHNO") & "$" & SELDREQRS("REQNO")
                        SELDREQRS.MoveNext
                    Wend
                    
                    'cmb_preinvest.ListIndex = 0
                End If

  ' FILLING THE PREVIOUSLY INVESTIGATED REQ NO.
                For I = 0 To 20
                    cmb_cnoslides.AddItem I
                    'cmb_cyoutspecimen.AddItem i
                    cmb_coutslides.AddItem I
                Next I
End Sub



Public Sub FACLEAR()
CMB_FNPROCEDURE.ListIndex = -1
cmb_fnprocode.ListIndex = -1
cmb_fnsite.Text = ""
cmb_fnlaterality1.Text = ""
txt_fnmetastasis.Text = ""
cmb_fnlaterality.Text = ""
TXT_FNSLIDEDRY.Text = ""
cmb_fnac.Text = ""
cmb_outsideslidesfn.Text = ""
cmb_fnpreinvest.Text = ""
txt_fnnmeterial.Text = ""
'CMB_CELLBLOCK.ListIndex = -1
End Sub

Public Sub FAASSIGN()
Dim durill As String
Dim I As Integer
On Error GoTo MSG:
'*** INSERTING INT0 RREQUISITION
INSRREQUISITN_QR(0) = req_chglen(Trim(txt_FNreqno.Text))
INSRREQUISITN_QR(1) = GET_REVDATE(CStr(DATE11))
INSRREQUISITN_QR(2) = Trim(txt_time.Text)
INSRREQUISITN_QR(3) = Trim(Msk_caseno.FormattedText)
INSRREQUISITN_QR(4) = Trim(txt_dept.Text)
INSRREQUISITN_QR(5) = Trim(txt_modcode.Text)
INSRREQUISITN_QR(6) = Trim(txt_status.Text)
INSRREQUISITN_QR(7) = Trim(txt_catgory.Text)
INSRREQUISITN_QR(8) = Trim(cmb_fnprocode.Text)
INSRREQUISITN_QR(9) = UCase(Trim(cmb_refdoc.Text))
'INSRREQUISITN_QR(9) = Trim(txt_pridocser.Text)

'Added by Milind on 27-03-2006 for inserting date for the samples collected at outside
    If Check1.Value = 1 Then
        Dim Dt As String
        Dt = "18/02/1980"
        INSRREQUISITN_QR(15) = Format(CDate(Dt), "dd/mm/yyyy")
    Else
        INSRREQUISITN_QR(15) = Null
    End If
'upto here

    If SELRREQUISITNFLAG = False Then
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
        SELRREQUISITNFLAG = True
    Else
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        SELRREQUISITN_RS.MoreResults
        SELRREQUISITN_RS.Requery
    End If
    If SELRREQUISITN_RS.RowCount > 0 Then
        INSRREQUISITN_QR(10) = "YES"
    Else
        INSRREQUISITN_QR(10) = "NO"
    End If
If opt_patient.Value = True Then
    If Trim(cmb_fnprocode.Text) = "F401" Then
        NO_CHARGEFA
    Else
        NO_CHARGEFA1
    End If
    If No_Chargeflag = True Then
        INSRREQUISITN_QR(11) = "NOCHARGE"
        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
        INSRREQUISITN_QR(13) = "P"
        No_Chargeflag = False
    ''ICU
    ElseIf lbl_icureq.Visible = True Then
        'If Trim(CMB_ICU1.Text) = "YES" Then
            INSRREQUISITN_QR(11) = Null
            INSRREQUISITN_QR(12) = Null
            'INSRREQUISITN_QR(13) = "P"
           ' INSRREQUISITN_QR(13) = "E"
         INSRREQUISITN_QR(13) = Null
        'End If
    ''ICU
    Else
        If Trim(CMB_FANIP.Text) = "LR" Or Trim(CMB_FANIP.Text) = "BAR" Or Trim(CMB_FANIP.Text) = "DrOp" Then
            
        If Trim(CMB_FANIP.Text) = "LR" Then
            INSRREQUISITN_QR(11) = "LR"
            INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
            INSRREQUISITN_QR(13) = "P"
        End If
        If Trim(CMB_FANIP.Text) = "BAR" Then
            INSRREQUISITN_QR(11) = "BAR"
            INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
            INSRREQUISITN_QR(13) = "P"
        End If
        If Trim(CMB_FANIP.Text) = "DrOp" Then
            INSRREQUISITN_QR(11) = "DrOp"
            INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
            INSRREQUISITN_QR(13) = "P"
        End If
    Else
        INSRREQUISITN_QR(11) = Null
        INSRREQUISITN_QR(12) = Null
        INSRREQUISITN_QR(13) = Null
End If
    End If
End If
If opt_staff.Value = True Then
        INSRREQUISITN_QR(11) = "NOCHARGE"
        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
        INSRREQUISITN_QR(13) = "P"
End If
If Trim(CMB_FANIP.Text) = "NONE" Or Len(Trim(CMB_FANIP.Text)) = 0 Then
    INSRREQUISITN_QR(14) = 0
End If
If Trim(CMB_FANIP.Text) = "VPP" Then
    INSRREQUISITN_QR(14) = 1
End If
If Trim(CMB_FANIP.Text) = "LR" Then
    INSRREQUISITN_QR(14) = 2
End If
If Trim(CMB_FANIP.Text) = "BAR" Then
    INSRREQUISITN_QR(14) = 3
End If
If Trim(CMB_FANIP.Text) = "DrOp" Then
    INSRREQUISITN_QR(14) = 4
End If
INSRREQUISITN_QR.Execute
    '*** INSERTING INT0 RREQUISITION
    '**** INSERTING INTO DREQUISITION

INSDREQUISITN_QR(0) = req_chglen(Trim(txt_FNreqno.Text))
INSDREQUISITN_QR(1) = ""
INSDREQUISITN_QR(2) = Trim(Msk_caseno.FormattedText)
                            
 I = InStr(cmb_fnsite, "$")
 If I = 0 Then
                            
               If sitedupflag = False Then
                    sitedupqr(0) = UCase(Trim(cmb_fnsite.Text))
                    Set siteduprs = sitedupqr.OpenResultset(1, 3)
                    sitedupflag = True
               Else
                    sitedupqr(0) = UCase(Trim(cmb_fnsite.Text))
                    siteduprs.MoreResults
                    siteduprs.Requery
                                   
               End If
                            
        If siteduprs.RowCount > 0 Then
                INSDREQUISITN_QR(3) = siteduprs("SITECD")
        Else
                If sitedup1flag = False Then
                    'sitedupqr(0) = Trim(cmb_sites.Text)
                    Set siteduprs1 = sitedupqr1.OpenResultset(1, 3)
                    sitedup1flag = True
                Else
                    siteduprs1.MoreResults
                    siteduprs1.Requery
                End If
                If (siteduprs1.RowCount > 0) Then
                    Dim a As Integer
                    If Len(Trim(cmb_fnsite.Text)) > 0 Then
                                     
                    a = siteduprs1(0)
                    a = a + 1
                    siteinsertqr(0) = a
                    siteinsertqr(1) = UCase(Trim(cmb_fnsite.Text))
                    siteinsertqr.Execute
                    INSDREQUISITN_QR(3) = a
                    Else
                        INSDREQUISITN_QR(3) = "9999"
                    End If
                 Else
                    
                    a = 5000
                    siteinsertqr(0) = a
                    siteinsertqr(1) = UCase(Trim(cmb_fnsite.Text))
                    siteinsertqr.Execute
                    INSDREQUISITN_QR(3) = a
                 End If
         End If
  Else
        INSDREQUISITN_QR(3) = Mid(cmb_fnsite, I + 1, 4)
  End If
INSDREQUISITN_QR(4) = Trim(cmb_fnlaterality.Text)
dur_ill = txt_year & "/" & Trim(CMB_DURILLMONTH.Text)
INSDREQUISITN_QR(5) = dur_ill
INSDREQUISITN_QR(6) = Trim(txt_fnmetastasis.Text)
INSDREQUISITN_QR(7) = Trim(cmb_fnlaterality1.Text)
INSDREQUISITN_QR(8) = Trim(TXT_FNSLIDEDRY.Text)
INSDREQUISITN_QR(9) = Trim(TXT_FNOSLIDEWET.Text)
INSDREQUISITN_QR(10) = Trim(cmb_outsideslidesfn.Text)
INSDREQUISITN_QR(11) = Null
If cmb_frozensec.Text = "YES" Then
    INSDREQUISITN_QR(12) = GET_REVDATE(Trim(MSK_FROSECDT.FormattedText))
Else
    INSDREQUISITN_QR(12) = Null
End If
INSDREQUISITN_QR(13) = Null
INSDREQUISITN_QR(14) = UCase(Trim(cmb_refdoc.Text))
INSDREQUISITN_QR(15) = Trim(cmb_reftodoc.Text)
INSDREQUISITN_QR(16) = Trim(TXT_OPPROCEDURE.Text)
If Trim(txt_status.Text) = "IP" Then
    INSDREQUISITN_QR(17) = Trim(txt_ward.Text)
    INSDREQUISITN_QR(18) = Trim(txt_bedno.Text)
Else
    INSDREQUISITN_QR(17) = Null
    INSDREQUISITN_QR(18) = Null
End If
INSDREQUISITN_QR(19) = " "
If Len(Trim(txt_fnnmeterial.Text)) > 0 Then
    INSDREQUISITN_QR(20) = Null
    INSDREQUISITN_QR(20) = Trim(txt_fnnmeterial.Text)
Else
    INSDREQUISITN_QR(20) = Null
    INSDREQUISITN_QR(20) = " "
End If
INSDREQUISITN_QR(21) = Null
INSDREQUISITN_QR.Execute
'**** INSERTING INTO DREQUISITION

 '*** UPDATING THE PROVISIONAL DIAGNOSIS
 
If opt_patient.Value = True Then
If SELTA1012FLAG = False Then
    selta1012_qr(0) = Msk_caseno.FormattedText
    Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
    SELTA1012FLAG = True
Else
     selta1012_qr(0) = Msk_caseno.FormattedText
     selta1012_rs.MoreResults
     selta1012_rs.Requery
End If
If selta1012_rs.RowCount > 0 Then
    UPDTA1012_QR(0) = Trim(txt_indctinvstg.Text)
    UPDTA1012_QR(1) = Msk_caseno.FormattedText
    UPDTA1012_QR.Execute
Else
    insta1012_qr(0) = Msk_caseno.FormattedText
    insta1012_qr(1) = Trim(txt_indctinvstg.Text)
    insta1012_qr.Execute
End If
End If
'*** UPDATING THE PROVISIONAL DIAGNOSIS
'*** INSERTING INTO DMULPROCE FOR COLLECTION AND DEPT COLLECTION SCREEN
''    INSDMULPROCE_QR(0) = req_chglen(txt_spreqno.Text)
''    INSDMULPROCE_QR(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
''    INSDMULPROCE_QR(2) = Msk_caseno.FormattedText
''    INSDMULPROCE_QR(3) = cmb_procedurecode.Text
''    INSDMULPROCE_QR.Execute

'***INSERTING INTO DMULPROCE FOR COLLECTION AND DEPT COLLECTION SCREEN
'***INSERTING INTO DREQSTAFFD TABLE
If opt_staff.Value = True Then
    INSDREQSTAFFD_QR(0) = req_chglen(txt_FNreqno.Text)
    INSDREQSTAFFD_QR(1) = Msk_caseno.Text
    INSDREQSTAFFD_QR(2) = Trim(txt_name.Text)
    INSDREQSTAFFD_QR(3) = Trim(txt_Age)
    INSDREQSTAFFD_QR(4) = Trim(txt_sex)
    If Len(Trim(cmb_dependents.Text)) = 0 Then
        INSDREQSTAFFD_QR(5) = "SELF"
    Else
        I = InStr(cmb_dependents, "$")
        INSDREQSTAFFD_QR(5) = Mid(cmb_dependents, I + 1, 10)
    End If
    INSDREQSTAFFD_QR.Execute
End If
'***INSERTING INTO DREQSTAFFD TABLE
    If SELCLINHISTFLAG = False Then
        SELCLINHIST_QR(0) = Msk_caseno.FormattedText
        Set SELCLINHIST_RS = SELCLINHIST_QR.OpenResultset(1, 3)
        SELCLINHISTFLAG = True
    Else
        SELCLINHIST_QR(0) = Msk_caseno.FormattedText
        SELCLINHIST_RS.MoreResults
        SELCLINHIST_RS.Requery
    End If
'    If SELCLINHIST_RS.RowCount > 0 Then
'        updclinhist_qr(0) = Trim(txt_clinhist.Text)
'        updclinhist_qr(1) = Trim(Msk_caseno.Text)
'        updclinhist_qr.Execute
'    Else
        insclinhist_qr(0) = Trim(Msk_caseno.FormattedText)
        insclinhist_qr(1) = GET_REVDATE(Trim(DATE11))
        insclinhist_qr(2) = Trim(time1)
        If Trim(txt_clinhist.Text) = "NYD" Then
            insclinhist_qr(3) = "NYD"
        Else
            insclinhist_qr(3) = Trim(txt_clinhist.Text)
        End If
        insclinhist_qr.Execute
    'End If
    FNACREQNO = txt_FNreqno.Text
    
''UPDATE DREQUISITN TABLE
If SELDISICU_FLAG = False Then
    SELDISICU_QR(0) = Trim(ipadd)
    Set SELDISICU_RS = SELDISICU_QR.OpenResultset(1, 3)
    SELDISICU_FLAG = True
Else
    SELDISICU_QR(0) = Trim(ipadd)
    SELDISICU_RS.MoreResults
    SELDISICU_RS.Requery
End If
If Not SELDISICU_RS.EOF Then
    UPDICU_QR(0) = "YES"
    UPDICU_QR(1) = req_chglen(Trim(txt_FNreqno.Text))
    UPDICU_QR.Execute
End If
''UPDATE DREQUISITN TABLE

Data1.UpdateRecord
Data1.Recordset.Bookmark = Data1.Recordset.LastModified
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Public Sub NO_CHARGEFA()
On Error GoTo MSG:
No_Chargeflag = False

If Left(Msk_caseno.FormattedText, 1) = "P" And Trim(txt_dept.Text) = "FZZ" And Trim(txt_modcode.Text) = "FA" Then
  'pre_sercodech(0) = Trim(cbo_proccode.Text)
  'Set selma101_RS = pre_sercodech.OpenResultset(1, 1)
      'If selma101_RS.RowCount = 1 And selma101_RS("yn_flag") = "Y" Then
          PRE_QUERY(0) = Msk_caseno.FormattedText
          'PRE_QUERY(1) = selma101_RS("FROM_DATE")
          'PRE_QUERY(2) = selma101_RS("TO_DATE")
          Set PREV_RS = PRE_QUERY.OpenResultset(1, 3)
            If PREV_RS.RowCount > 0 And Trim(txt_dept.Text) = "FZZ" Then
                   No_Chargeflag = True
            Else
                  If Trim(txt_catgory.Text) = "NC" Then
                    No_Chargeflag = True
                  Else
                        No_Chargeflag = False
                        NoMemo = False
                     '   MsgBox "Please Prepare the Memo !!", vbInformation, "Requisition"
                  End If
            End If
            'End If
     
Else
  'SELMA101_QR(0) = Trim(cbo_proccode.Text)
  'SELMA101_RS.MoreResults
  'SELMA101_RS.Requery
If SELMA101FLAG = False Then
    SELMA101_QR(0) = Trim(cmb_fnprocode.Text)
    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
    SELMA101FLAG = True
 Else
    SELMA101_QR(0) = Trim(cmb_fnprocode.Text)
    SELMA101_RS.MoreResults
    SELMA101_RS.Requery
End If

'SELMA101_QR(0) = Trim(cmb_fnprocode.Text)
'SELMA101_RS.MoreResults
'SELMA101_RS.Requery

If SELMA101_RS.RowCount > 0 Then
    If Trim(txt_catgory.Text) = "C" Then
        If SELMA101_RS("C_CH") = 0 Then
            No_Chargeflag = True
        End If
    ElseIf Trim(txt_catgory.Text) = "NC" Then
        If SELMA101_RS("NC_CH") = 0 Then
            No_Chargeflag = True
        End If
    End If
    If Trim(txt_catgory.Text) = "B" Then
        If SELMA101_RS("B_HOSP_CH") = 0 Then
            No_Chargeflag = True
        End If
    End If
Else
    'MsgBox "There is NO  RECORD with this SERVICE CODE."  'COMMENTED BY SHWETA 18-AUG-09
End If
End If

If No_Chargeflag = False Then
       NOCHARG_QR(0) = Msk_caseno.FormattedText
        If NOCHARG_FLG = False Then
          Set NOCHARG_RS = NOCHARG_QR.OpenResultset(2, 3)
          NOCHARG_FLG = True
        Else
          NOCHARG_RS.MoreResults
          NOCHARG_RS.Requery
        End If
        If NOCHARG_RS.RowCount > 0 Then
          'CHCKIN FREE CODE
            PRFREE_QR(0) = Trim(cmb_fnprocode.Text)
              If PRFREE_FLG = False Then
                 Set PRFREE_RS = PRFREE_QR.OpenResultset(2, 3)
                 PRFREE_FLG = True
              Else
                 PRFREE_RS.MoreResults
                 PRFREE_RS.Requery
              End If
              If PRFREE_RS.RowCount > 0 Then
                 No_Chargeflag = True
                 NoMemo = True
              End If
        Else
            No_Chargeflag = False
            NoMemo = False
            MsgBox "Please Prepare the Memo !!", vbInformation, "Requisition"
        End If

End If


Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Public Sub FAENABL(str As String)
    If str = "ENA" Then
        CMB_FNPROCEDURE.Enabled = True
        cmb_fnprocode.Enabled = True
        cmb_fnsite.Enabled = True
        cmb_fnlaterality.Enabled = True
        cmb_fnlaterality1.Enabled = True
        txt_fnmetastasis.Enabled = True
        TXT_FNSLIDEDRY.Enabled = True
        TXT_FNOSLIDEWET.Enabled = True
        cmb_fnac.Enabled = True
        cmb_outsideslidesfn.Enabled = True
        cmb_fnpreinvest.Enabled = True
        'CMB_CELLBLOCK.Enabled = True
        txt_fnnmeterial.Enabled = True
    End If
    If str = "DIS" Then
        CMB_FNPROCEDURE.Enabled = False
        cmb_fnprocode.Enabled = False
        cmb_fnsite.Enabled = False
        cmb_fnlaterality.Enabled = False
        cmb_fnlaterality1.Enabled = False
        txt_fnmetastasis.Enabled = False
        TXT_FNSLIDEDRY.Enabled = False
        TXT_FNOSLIDEWET.Enabled = False
        cmb_fnac.Enabled = False
        cmb_outsideslidesfn.Enabled = False
        cmb_fnpreinvest.Enabled = False
        'CMB_CELLBLOCK.Enabled = False
        txt_fnnmeterial.Enabled = False
    End If
    If str = "MOD" Then
        CMB_FNPROCEDURE.Enabled = False
        cmb_fnprocode.Enabled = False
        cmb_fnsite.Enabled = True
        cmb_fnlaterality.Enabled = True
        cmb_fnlaterality1.Enabled = True
        txt_fnmetastasis.Enabled = True
        TXT_FNSLIDEDRY.Enabled = True
        TXT_FNOSLIDEWET.Enabled = False
        cmb_fnac.Enabled = True
        cmb_outsideslidesfn.Enabled = False
        cmb_fnpreinvest.Enabled = True
        'CMB_CELLBLOCK.Enabled = True
        txt_fnnmeterial.Enabled = True
    End If
End Sub

Public Sub FAUPD()
        '**** UPDATING THE REQUISITION TABLE
        updrrequisitn_qr(0) = cmb_refdoc.Text
        updrrequisitn_qr(1) = req_chglen(cmb_reqno.Text)
        updrrequisitn_qr.Execute
        '**** UPDATING THE REQUISITION TABLE
        
        '**** Updating the drequisition table
        upddrequisitn_qr(0) = UCase(cmb_refdoc.Text)
        upddrequisitn_qr(1) = UCase(cmb_reftodoc.Text)
        upddrequisitn_qr(2) = txt_year & "/" & CMB_DURILLMONTH
        I = InStr(cmb_sites, "$")
        If I = 0 Then
               If sitedupflag = False Then
                    sitedupqr(0) = UCase(Trim(cmb_sites.Text))
                    Set siteduprs = sitedupqr.OpenResultset(1, 3)
                    sitedupflag = True
               Else
                    sitedupqr(0) = UCase(Trim(cmb_sites.Text))
                    siteduprs.MoreResults
                    siteduprs.Requery
                                   
               End If
               If siteduprs.RowCount > 0 Then
                    upddrequisitn_qr(3) = siteduprs("SITECD")
                          
               Else
               
                    If sitedup1flag = False Then
                        'sitedupqr(0) = Trim(cmb_sites.Text)
                        Set siteduprs1 = sitedupqr1.OpenResultset(1, 3)
                        sitedup1flag = True
                    Else
                        siteduprs1.MoreResults
                        siteduprs1.Requery
                    End If
                    If (siteduprs1.RowCount > 0) Then
                        
                        Dim a As Integer
                        If Len(Trim(cmb_sites.Text)) > 0 Then
                            a = siteduprs1(0)
                            a = a + 1
                            siteinsertqr(0) = a
                            siteinsertqr(1) = UCase(Trim(cmb_sites.Text))
                            siteinsertqr.Execute
                            upddrequisitn_qr(3) = a
                         Else
                            upddrequisitn_qr(3) = "9999"
                         End If
                        
                     Else
                        
                        a = 5000
                        siteinsertqr(0) = a
                        siteinsertqr(1) = UCase(Trim(cmb_sites.Text))
                        siteinsertqr.Execute
                        upddrequisitn_qr(3) = a
                     End If
                End If
         Else
               upddrequisitn_qr(3) = Mid(cmb_sites, I + 1, 4)
         End If
          
        
        upddrequisitn_qr(4) = UCase(Trim(cmb_fnlaterality))
        upddrequisitn_qr(5) = UCase(Trim(cmb_fnac))
        If Len(Trim(txt_fnnmeterial.Text)) > 0 Then
            upddrequisitn_qr(6) = Null
            upddrequisitn_qr(6) = UCase(Trim(txt_fnnmeterial.Text))
        Else
            upddrequisitn_qr(6) = Null
            upddrequisitn_qr(6) = " "
        End If
        upddrequisitn_qr(7) = UCase(Trim(txt_fnmetastasis.Text))
        upddrequisitn_qr(8) = UCase(Trim(cmb_fnlaterality1.Text))
        upddrequisitn_qr(9) = cmb_outsideslidesfn
        upddrequisitn_qr(10) = txt_outspecimen.Text
        upddrequisitn_qr(11) = TXT_FNSLIDEDRY
        upddrequisitn_qr(12) = "NO"
        upddrequisitn_qr(13) = Null
        upddrequisitn_qr(14) = Null
        upddrequisitn_qr(15) = req_chglen(cmb_reqno.Text)
        upddrequisitn_qr.Execute
        '**** Updating the drequisition table
         '*** UPDATING THE PROVISIONAL DIAGNOSIS
        If opt_patient.Value = True Then
        If SELTA1012FLAG = False Then
            selta1012_qr(0) = Msk_caseno.FormattedText
            Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
            SELTA1012FLAG = True
        Else
             selta1012_qr(0) = Msk_caseno.FormattedText
             selta1012_rs.MoreResults
             selta1012_rs.Requery
        End If
        If selta1012_rs.RowCount > 0 Then
            UPDTA1012_QR(0) = Trim(txt_indctinvstg.Text)
            UPDTA1012_QR(1) = Msk_caseno.FormattedText
            UPDTA1012_QR.Execute
        Else
            insta1012_qr(0) = Msk_caseno.FormattedText
            insta1012_qr(1) = Trim(txt_indctinvstg.Text)
            insta1012_qr.Execute
        End If
        End If
        '*** UPDATING THE PROVISIONAL DIAGNOSIS
        '*** updating the clinical history
        If SELCLINHISTFLAG = False Then
            SELCLINHIST_QR(0) = Msk_caseno.FormattedText
            Set SELCLINHIST_RS = SELCLINHIST_QR.OpenResultset(1, 3)
            SELCLINHISTFLAG = True
        Else
            SELCLINHIST_QR(0) = Msk_caseno.FormattedText
            SELCLINHIST_RS.MoreResults
            SELCLINHIST_RS.Requery
        End If
'        If SELCLINHIST_RS.RowCount > 0 Then
'            updclinhist_qr(0) = Trim(txt_clinhist.Text)
'            updclinhist_qr(1) = Trim(Msk_caseno.Text)
'            updclinhist_qr.Execute
'        Else
            insclinhist_qr(0) = Trim(Msk_caseno.FormattedText)
            insclinhist_qr(1) = GET_REVDATE(Trim(DATE11))
            insclinhist_qr(2) = Trim(time1)
            If Trim(txt_clinhist.Text) = "NYD" Then
                insclinhist_qr(3) = "NYD"
            Else
                insclinhist_qr(3) = Trim(txt_clinhist.Text)
            End If
            insclinhist_qr.Execute
        'End If
        
        If SELRREQUISITNFLAG2 = False Then
                SELRREQUISITN_QR2(0) = req_chglen(Trim(cmb_reqno.Text))
                Set SELRREQUISITN_RS2 = SELRREQUISITN_QR2.OpenResultset(1, 3)
                SELRREQUISITNFLAG2 = True
           Else
                SELRREQUISITN_QR2(0) = req_chglen(Trim(cmb_reqno.Text))
                SELRREQUISITN_RS2.MoreResults
                SELRREQUISITN_RS2.Requery
           End If
        If SELRREQUISITN_RS2.RowCount > 0 Then
                If IsNull(SELRREQUISITN_RS2("reqnstat")) Then
        
                If CMB_FANIP.Text = "VPP" Or CMB_FANIP.Text = "LR" Or CMB_FANIP.Text = "BAR" Or CMB_FANIP.Text = "DrOp" Then
                    If CMB_FANIP.Text = "VPP" Then
                        UPDRREQQR(0) = Null
                        UPDRREQQR(1) = Null
                        UPDRREQQR(2) = Null
                        UPDRREQQR(3) = 1
                        UPDRREQQR(4) = req_chglen(cmb_reqno.Text)
                        UPDRREQQR.Execute
                    End If
                    If CMB_FANIP.Text = "LR" Then
                        UPDRREQQR(0) = "LR"
                        UPDRREQQR(1) = GET_REVDATE(Trim(DATE11))
                        UPDRREQQR(2) = "P"
                        UPDRREQQR(3) = 2
                        UPDRREQQR(4) = req_chglen(cmb_reqno.Text)
                        UPDRREQQR.Execute
                    End If
                    If CMB_FANIP.Text = "BAR" Then
                        UPDRREQQR(0) = "BAR"
                        UPDRREQQR(1) = GET_REVDATE(Trim(DATE11))
                        UPDRREQQR(2) = "P"
                        UPDRREQQR(3) = 3
                        UPDRREQQR(4) = req_chglen(cmb_reqno.Text)
                        UPDRREQQR.Execute
                    End If
                    If CMB_FANIP.Text = "DrOp" Then
                        UPDRREQQR(0) = "DrOp"
                        UPDRREQQR(1) = GET_REVDATE(Trim(DATE11))
                        UPDRREQQR(2) = "P"
                        UPDRREQQR(3) = 4
                        UPDRREQQR(4) = req_chglen(cmb_reqno.Text)
                        UPDRREQQR.Execute
                    End If
                End If
              End If
       End If
        
        
End Sub

Public Sub CGCLEAR()
    CMB_CGPROCEDURE.ListIndex = -1
    cbo_proccode.ListIndex = -1
    cmb_cnsite.Text = ""
    cmb_claterality.Text = ""
    txt_trtdetail.Text = ""
    txt_cnatmeterial.Text = ""
    cmb_cnoslides.Text = ""
    cmb_cyoutspecimen.Text = ""
    cmb_coutslides.Text = ""
    txt_scopyfind.Text = ""
    MSK_LMP.Mask = "##/##/####"
    MSK_LMP.Text = "__/__/____"
    cmb_cpreinv.Text = ""
End Sub

Public Sub CGASSIGN()
Dim durill As String
Dim I As Integer
On Error GoTo MSG:
'*** INSERTING INT0 RREQUISITION
INSRREQUISITN_QR(0) = req_chglen(Trim(txt_cyreqno1.Text))
INSRREQUISITN_QR(1) = GET_REVDATE(CStr(DATE11))
INSRREQUISITN_QR(2) = Trim(txt_time.Text)
INSRREQUISITN_QR(3) = Trim(Msk_caseno.FormattedText)
INSRREQUISITN_QR(4) = Trim(txt_dept.Text)
INSRREQUISITN_QR(5) = Trim(txt_modcode.Text)
INSRREQUISITN_QR(6) = Trim(txt_status.Text)
INSRREQUISITN_QR(7) = Trim(txt_catgory.Text)
INSRREQUISITN_QR(8) = Trim(cbo_proccode.Text)
INSRREQUISITN_QR(9) = UCase(Trim(cmb_refdoc.Text))
'INSRREQUISITN_QR(9) = Trim(txt_pridocser.Text)

'Added by Milind on 27-03-2006 for inserting date for the samples collected at outside
    If Check1.Value = 1 Then
        Dim Dt As String
        Dt = "18/02/1980"
        INSRREQUISITN_QR(15) = Format(CDate(Dt), "dd/mm/yyyy")
    Else
        INSRREQUISITN_QR(15) = Null
    End If
'upto here

    If SELRREQUISITNFLAG = False Then
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
        SELRREQUISITNFLAG = True
    Else
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        SELRREQUISITN_RS.MoreResults
        SELRREQUISITN_RS.Requery
    End If
    If SELRREQUISITN_RS.RowCount > 0 Then
        INSRREQUISITN_QR(10) = "YES"
    Else
        INSRREQUISITN_QR(10) = "NO"
    End If
If opt_patient.Value = True Then
    NO_CHARGECG
    If No_Chargeflag = True Then
        INSRREQUISITN_QR(11) = "NOCHARGE"
        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
        INSRREQUISITN_QR(13) = "P"
        No_Chargeflag = False
    ''ICU
    ElseIf lbl_icureq.Visible = True Then
        'If Trim(CMB_ICU1.Text) = "YES" Then
            INSRREQUISITN_QR(11) = Null
            INSRREQUISITN_QR(12) = Null
           ' INSRREQUISITN_QR(13) = "P"
           ' INSRREQUISITN_QR(13) = "E"
           'INSRREQUISITN_QR(13) = " "
            INSRREQUISITN_QR(13) = Null
        'End If
    ''ICU
    Else
        If Trim(CMB_CGNIP.Text) = "LR" Or Trim(CMB_CGNIP.Text) = "BAR" Or Trim(CMB_CGNIP.Text) = "DrOp" Then
            
            If Trim(CMB_CGNIP.Text) = "LR" Then
                INSRREQUISITN_QR(11) = "LR"
                INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
                INSRREQUISITN_QR(13) = "P"
            End If
            If Trim(CMB_CGNIP.Text) = "BAR" Then
                INSRREQUISITN_QR(11) = "BAR"
                INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
                INSRREQUISITN_QR(13) = "P"
            End If
            If Trim(CMB_CGNIP.Text) = "DrOp" Then
                INSRREQUISITN_QR(11) = "DrOp"
                INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
                INSRREQUISITN_QR(13) = "P"
            End If
        Else
            INSRREQUISITN_QR(11) = Null
            INSRREQUISITN_QR(12) = Null
            INSRREQUISITN_QR(13) = Null
        End If
    End If
 
End If
If opt_staff.Value = True Then
        INSRREQUISITN_QR(11) = "NOCHARGE"
        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
        INSRREQUISITN_QR(13) = "P"
End If

   If Trim(cmb_vpp.Text) = "NONE" Or Len(Trim(cmb_vpp.Text)) = 0 Or Trim(cmb_vpp.Text) = "" Then
    INSRREQUISITN_QR(14) = 0
End If
If Trim(cmb_vpp.Text) = "VPP" Then
    INSRREQUISITN_QR(14) = 1
End If
If Trim(cmb_vpp.Text) = "DER" Then
    INSRREQUISITN_QR(14) = 2
End If
If Trim(cmb_vpp.Text) = "BAR" Then
    INSRREQUISITN_QR(14) = 3
End If
If Trim(cmb_vpp.Text) = "DrOp" Then
    INSRREQUISITN_QR(14) = 4
End If

INSRREQUISITN_QR.Execute
    '*** INSERTING INT0 RREQUISITION
    '**** INSERTING INTO DREQUISITION

INSDREQUISITN_QR(0) = req_chglen(Trim(txt_cyreqno1.Text))
INSDREQUISITN_QR(1) = Trim(txt_sppathno.Text)
INSDREQUISITN_QR(2) = Trim(Msk_caseno.FormattedText)
                            
 I = InStr(cmb_cnsite, "$")
 If I = 0 Then
                            
               If sitedupflag = False Then
                    sitedupqr(0) = UCase(Trim(cmb_cnsite.Text))
                    Set siteduprs = sitedupqr.OpenResultset(1, 3)
                    sitedupflag = True
               Else
                    sitedupqr(0) = UCase(Trim(cmb_cnsite.Text))
                    siteduprs.MoreResults
                    siteduprs.Requery
                                   
               End If
                            
        If siteduprs.RowCount > 0 Then
                INSDREQUISITN_QR(3) = siteduprs("SITECD")
        Else
                If sitedup1flag = False Then
                    'sitedupqr(0) = Trim(cmb_sites.Text)
                    Set siteduprs1 = sitedupqr1.OpenResultset(1, 3)
                    sitedup1flag = True
                Else
                    siteduprs1.MoreResults
                    siteduprs1.Requery
                End If
                If (siteduprs1.RowCount > 0) Then
                    Dim a As Integer
                    If Len(Trim(cmb_cnsite.Text)) > 0 Then
                                     
                    a = siteduprs1(0)
                    a = a + 1
                    siteinsertqr(0) = a
                    siteinsertqr(1) = UCase(Trim(cmb_cnsite.Text))
                    siteinsertqr.Execute
                    INSDREQUISITN_QR(3) = a
                    Else
                        INSDREQUISITN_QR(3) = "9999"
                    End If
                 Else
                    
                    a = 5000
                    siteinsertqr(0) = a
                    siteinsertqr(1) = UCase(Trim(cmb_cnsite.Text))
                    siteinsertqr.Execute
                    INSDREQUISITN_QR(3) = a
                 End If
         End If
  Else
        INSDREQUISITN_QR(3) = Mid(cmb_cnsite, I + 1, 4)
  End If
INSDREQUISITN_QR(4) = Trim(cmb_claterality.Text)
dur_ill = txt_year & "/" & Trim(CMB_DURILLMONTH.Text)
INSDREQUISITN_QR(5) = dur_ill
INSDREQUISITN_QR(6) = Trim(txt_cmetastasis.Text)
INSDREQUISITN_QR(7) = Null
INSDREQUISITN_QR(8) = Null
INSDREQUISITN_QR(9) = Trim(cmb_cnoslides.Text)
INSDREQUISITN_QR(10) = Trim(cmb_coutslides.Text)
INSDREQUISITN_QR(11) = ""
If Trim(MSK_LMP.Text) <> "" Then
    INSDREQUISITN_QR(12) = GET_REVDATE(Trim(MSK_LMP.FormattedText))
Else
    INSDREQUISITN_QR(12) = Null
End If
INSDREQUISITN_QR(13) = Null
INSDREQUISITN_QR(14) = Trim(cmb_refdoc.Text)
INSDREQUISITN_QR(15) = Trim(cmb_reftodoc.Text)
INSDREQUISITN_QR(16) = Null
If Trim(txt_status.Text) = "IP" Then
    INSDREQUISITN_QR(17) = Trim(txt_ward.Text)
    INSDREQUISITN_QR(18) = Trim(txt_bedno.Text)
Else
    INSDREQUISITN_QR(17) = Null
    INSDREQUISITN_QR(18) = Null
End If
INSDREQUISITN_QR(19) = " "
If Len(Trim(txt_cnatmeterial.Text)) > 0 Then
    INSDREQUISITN_QR(20) = Null
    INSDREQUISITN_QR(20) = txt_cnatmeterial.Text
Else
    INSDREQUISITN_QR(20) = Null
    INSDREQUISITN_QR(20) = " "
End If
INSDREQUISITN_QR(21) = Null
INSDREQUISITN_QR.Execute
'**** INSERTING INTO DREQUISITION
'**** INSERTING INTO DSFTDCY SCOPY FINDING AND TREATMENT DETAILS
insdsftdcyqr(0) = req_chglen(txt_cyreqno1.Text)
If Len(Trim(txt_scopyfind.Text)) > 0 Then
    insdsftdcyqr(1) = Null
    insdsftdcyqr(1) = txt_scopyfind.Text
Else
    insdsftdcyqr(1) = Null
    insdsftdcyqr(1) = " "
End If
If Len(Trim(txt_scopyfind.Text)) > 0 Then
    insdsftdcyqr(2) = Null
    insdsftdcyqr(2) = txt_trtdetail.Text
Else
    insdsftdcyqr(2) = Null
    insdsftdcyqr(2) = " "
End If
If Len(Trim(cmb_cyoutspecimen.Text)) > 0 Then
    insdsftdcyqr(3) = Null
    insdsftdcyqr(3) = cmb_cyoutspecimen.Text
Else
    insdsftdcyqr(3) = Null
    insdsftdcyqr(3) = " "
End If
insdsftdcyqr.Execute

'**** INSERTING INTO DSFTDCY SCOPY FINDING AND TREATMENT DETAILS
 '*** UPDATING THE PROVISIONAL DIAGNOSIS
 
If opt_patient.Value = True Then
If SELTA1012FLAG = False Then
    selta1012_qr(0) = Msk_caseno.FormattedText
    Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
    SELTA1012FLAG = True
Else
     selta1012_qr(0) = Msk_caseno.FormattedText
     selta1012_rs.MoreResults
     selta1012_rs.Requery
End If
If selta1012_rs.RowCount > 0 Then
    UPDTA1012_QR(0) = Trim(txt_indctinvstg.Text)
    UPDTA1012_QR(1) = Msk_caseno.FormattedText
    UPDTA1012_QR.Execute
Else
    insta1012_qr(0) = Msk_caseno.FormattedText
    insta1012_qr(1) = Trim(txt_indctinvstg.Text)
    insta1012_qr.Execute
End If
End If
'*** UPDATING THE PROVISIONAL DIAGNOSIS
'*** INSERTING INTO DMULPROCE FOR COLLECTION AND DEPT COLLECTION SCREEN
''    INSDMULPROCE_QR(0) = req_chglen(txt_spreqno.Text)
''    INSDMULPROCE_QR(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
''    INSDMULPROCE_QR(2) = Msk_caseno.FormattedText
''    INSDMULPROCE_QR(3) = cmb_procedurecode.Text
''    INSDMULPROCE_QR.Execute

'***INSERTING INTO DMULPROCE FOR COLLECTION AND DEPT COLLECTION SCREEN
'***INSERTING INTO DREQSTAFFD TABLE
If opt_staff.Value = True Then
    INSDREQSTAFFD_QR(0) = req_chglen(txt_cyreqno1.Text)
    INSDREQSTAFFD_QR(1) = Msk_caseno.Text
    INSDREQSTAFFD_QR(2) = Trim(txt_name.Text)
    INSDREQSTAFFD_QR(3) = Trim(txt_Age)
    INSDREQSTAFFD_QR(4) = Trim(txt_sex)
    If Len(Trim(cmb_dependents.Text)) = 0 Then
        INSDREQSTAFFD_QR(5) = "SELF"
    Else
        I = InStr(cmb_dependents, "$")
        INSDREQSTAFFD_QR(5) = Mid(cmb_dependents, I + 1, 10)
    End If
    INSDREQSTAFFD_QR.Execute
End If
'***INSERTING INTO DREQSTAFFD TABLE
    If SELCLINHISTFLAG = False Then
        SELCLINHIST_QR(0) = Msk_caseno.FormattedText
        Set SELCLINHIST_RS = SELCLINHIST_QR.OpenResultset(1, 3)
        SELCLINHISTFLAG = True
    Else
        SELCLINHIST_QR(0) = Msk_caseno.FormattedText
        SELCLINHIST_RS.MoreResults
        SELCLINHIST_RS.Requery
    End If
'    If SELCLINHIST_RS.RowCount > 0 Then
'        updclinhist_qr(0) = Trim(txt_clinhist.Text)
'        updclinhist_qr(1) = Trim(Msk_caseno.Text)
'        updclinhist_qr.Execute
'    Else
        insclinhist_qr(0) = Trim(Msk_caseno.FormattedText)
        insclinhist_qr(1) = GET_REVDATE(Trim(DATE11))
        insclinhist_qr(2) = Trim(time1)
        If Trim(txt_clinhist.Text) = "NYD" Then
            insclinhist_qr(3) = "NYD"
        Else
            insclinhist_qr(3) = Trim(txt_clinhist.Text)
        End If
        insclinhist_qr.Execute
    'End If
''UPDATE DREQUISITN TABLE
If SELDISICU_FLAG = False Then
    SELDISICU_QR(0) = Trim(ipadd)
    Set SELDISICU_RS = SELDISICU_QR.OpenResultset(1, 3)
    SELDISICU_FLAG = True
Else
    SELDISICU_QR(0) = Trim(ipadd)
    SELDISICU_RS.MoreResults
    SELDISICU_RS.Requery
End If
If Not SELDISICU_RS.EOF Then
    UPDICU_QR(0) = "YES"
    UPDICU_QR(1) = req_chglen(Trim(txt_cyreqno1.Text))
    UPDICU_QR.Execute
End If
    
''UPDATE DREQUISITN TABLE
Data1.UpdateRecord
Data1.Recordset.Bookmark = Data1.Recordset.LastModified
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Public Sub NO_CHARGECG()

'*****
On Error GoTo MSG:
No_Chargeflag = False
 
If Left(Msk_caseno.FormattedText, 1) = "P" And Trim(txt_dept.Text) = "FZZ" Then
  'pre_sercodech(0) = Trim(cbo_proccode.Text)
  'Set selma101_RS = pre_sercodech.OpenResultset(1, 1)
      'If selma101_RS.RowCount = 1 And selma101_RS("yn_flag") = "Y" Then
          PRE_QUERY(0) = Msk_caseno.FormattedText
          'PRE_QUERY(1) = selma101_RS("FROM_DATE")
          'PRE_QUERY(2) = selma101_RS("TO_DATE")
          Set PREV_RS = PRE_QUERY.OpenResultset(1, 3)
            If PREV_RS.RowCount > 0 And Trim(txt_dept.Text) = "FZZ" Then
                   No_Chargeflag = True
            Else
                  If Trim(txt_catgory.Text) = "NC" Then
                    No_Chargeflag = True
                  Else
                    No_Chargeflag = True
                    NoMemo = True
                  End If
           End If
   
Else
    If SELMA101FLAG = False Then
        SELMA101_QR(0) = Trim(cbo_proccode.Text)
        Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
        SELMA101FLAG = True
     Else
        SELMA101_QR(0) = Trim(cbo_proccode.Text)
        SELMA101_RS.MoreResults
        SELMA101_RS.Requery
    End If

'  SELMA101_QR(0) = Trim(cbo_proccode.Text)
'  SELMA101_RS.MoreResults
'  SELMA101_RS.Requery

    If SELMA101_RS.RowCount > 0 Then
        If Trim(txt_catgory.Text) = "C" Then
            If SELMA101_RS("C_CH") = 0 Then
                No_Chargeflag = True
            End If
        ElseIf Trim(txt_catgory.Text) = "NC" Then ' Or Opt_S.Value = True Then
            If SELMA101_RS("NC_CH") = 0 Then
                No_Chargeflag = True
            End If
        ElseIf opt_staff.Value = True Then
                No_Chargeflag = True
        End If
    Else
        'MsgBox "There is NO  RECORD with this SERVICE CODE."  'COMMENTED BY SHWETA 18-AUG-09
    End If
End If

If No_Chargeflag = False Then
          NOCHARG_QR(0) = Msk_caseno.FormattedText
            If NOCHARG_FLG = False Then
              Set NOCHARG_RS = NOCHARG_QR.OpenResultset(2, 3)
              NOCHARG_FLG = True
            Else
              NOCHARG_RS.MoreResults
              NOCHARG_RS.Requery
            End If
            If NOCHARG_RS.RowCount > 0 Then
              'CHCKIN FREE CODE
                PRFREE_QR(0) = Trim(cmb_fnprocode.Text)
                  If PRFREE_FLG = False Then
                     Set PRFREE_RS = PRFREE_QR.OpenResultset(2, 3)
                     PRFREE_FLG = True
                  Else
                     PRFREE_RS.MoreResults
                     PRFREE_RS.Requery
                  End If
                  If PRFREE_RS.RowCount > 0 Then
                     No_Chargeflag = True
                     NoMemo = True
                  End If
           Else
            No_Chargeflag = False
            NoMemo = False
            MsgBox "Please Prepare the Memo !!", vbInformation, "Requisition"
           End If
End If

Exit Sub
MSG:
MsgBox ERROR(Err)



'*****


On Error GoTo MSG:
''No_Chargeflag = False
''
''SELMA101_QR(0) = Trim(cbo_proccode.Text)
''SELMA101_RS.MoreResults
''SELMA101_RS.Requery
''If SELMA101_RS.RowCount > 0 Then
''    If Trim(txt_catgory.Text) = "C" Then
''        If SELMA101_RS("C_CH") = 0 Then
''            No_Chargeflag = True
''        End If
''    ElseIf Trim(txt_catgory.Text) = "NC" Then
''        If SELMA101_RS("NC_CH") = 0 Then
''            No_Chargeflag = True
''        End If
''    End If
''Else
''    MsgBox "There is NO  RECORD with this SERVICE CODE."
''End If
''Exit Sub
''MSG:
''MsgBox Error(Err)
End Sub

Public Sub CGUPD()
'**** Updating the drequisition table
        upddrequisitn_qr(0) = (cmb_refdoc.Text)
        upddrequisitn_qr(1) = cmb_reftodoc.Text
        upddrequisitn_qr(2) = txt_year & "/" & CMB_DURILLMONTH
        I = InStr(cmb_cnsite, "$")
        If I = 0 Then
               If sitedupflag = False Then
                    sitedupqr(0) = UCase(Trim(cmb_cnsite.Text))
                    Set siteduprs = sitedupqr.OpenResultset(1, 3)
                    sitedupflag = True
               Else
                    sitedupqr(0) = UCase(Trim(cmb_cnsite.Text))
                    siteduprs.MoreResults
                    siteduprs.Requery
                                   
               End If
               If siteduprs.RowCount > 0 Then
                    upddrequisitn_qr(3) = siteduprs("SITECD")
                          
               Else
               
                    If sitedup1flag = False Then
                        'sitedupqr(0) = Trim(cmb_sites.Text)
                        Set siteduprs1 = sitedupqr1.OpenResultset(1, 3)
                        sitedup1flag = True
                    Else
                        siteduprs1.MoreResults
                        siteduprs1.Requery
                    End If
                    If (siteduprs1.RowCount > 0) Then
                        
                        Dim a As Integer
                        If Len(Trim(cmb_cnsite.Text)) > 0 Then
                            a = siteduprs1(0)
                            a = a + 1
                            siteinsertqr(0) = a
                            siteinsertqr(1) = UCase(Trim(cmb_cnsite.Text))
                            siteinsertqr.Execute
                            upddrequisitn_qr(3) = a
                         Else
                            upddrequisitn_qr(3) = "9999"
                         End If
                        
                     Else
                        
                        a = 5000
                        siteinsertqr(0) = a
                        siteinsertqr(1) = UCase(Trim(cmb_cnsite.Text))
                        siteinsertqr.Execute
                        upddrequisitn_qr(3) = a
                     End If
                End If
         Else
               upddrequisitn_qr(3) = Mid(cmb_cnsite, I + 1, 4)
         End If
              
        
        upddrequisitn_qr(4) = UCase(Trim(cmb_claterality))
        upddrequisitn_qr(5) = Null
        If Len(Trim(txt_cnatmeterial.Text)) > 0 Then
            upddrequisitn_qr(6) = Null
            upddrequisitn_qr(6) = UCase(Trim(txt_cnatmeterial.Text))
        Else
            upddrequisitn_qr(6) = Null
            upddrequisitn_qr(6) = " "
        End If
        upddrequisitn_qr(7) = UCase(Trim(txt_cmetastasis.Text))
        upddrequisitn_qr(8) = Null
        upddrequisitn_qr(9) = cmb_coutslides
        upddrequisitn_qr(10) = Null
        upddrequisitn_qr(11) = cmb_cnoslides.Text
        upddrequisitn_qr(12) = "NO"
        If txt_modcode.Text = "CG" Then
        If MSK_LMP.Text = "" Then
            upddrequisitn_qr(13) = Null
        Else
            upddrequisitn_qr(13) = GET_REVDATE(Trim(MSK_LMP.FormattedText))
        End If
        Else
         upddrequisitn_qr(13) = Null
        End If
        upddrequisitn_qr(14) = Null
        upddrequisitn_qr(15) = req_chglen(cmb_reqno.Text)
        upddrequisitn_qr.Execute
        '**** Updating the drequisition table
        '**** updating the scopyfinding and treatement details
        If Len(Trim(txt_scopyfind.Text)) > 0 Then
            upddsftdcyqr(0) = Null
            upddsftdcyqr(0) = txt_scopyfind.Text
        Else
            upddsftdcyqr(0) = Null
            upddsftdcyqr(0) = " "
        End If
        If Len(Trim(txt_trtdetail.Text)) > 0 Then
            upddsftdcyqr(1) = Null
            upddsftdcyqr(1) = txt_trtdetail.Text
        Else
            upddsftdcyqr(1) = Null
            upddsftdcyqr(1) = " "
        End If
        If Len(Trim(cmb_cyoutspecimen.Text)) > 0 Then
            upddsftdcyqr(2) = Null
            upddsftdcyqr(2) = cmb_cyoutspecimen.Text
        Else
            upddsftdcyqr(2) = Null
            upddsftdcyqr(2) = " "
        End If
            upddsftdcyqr(3) = req_chglen(cmb_reqno.Text)
            upddsftdcyqr.Execute
        '**** updating the scopyfinding and treatement details
         '*** UPDATING THE PROVISIONAL DIAGNOSIS
        If opt_patient.Value = True Then
        If SELTA1012FLAG = False Then
            selta1012_qr(0) = Msk_caseno.FormattedText
            Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
            SELTA1012FLAG = True
        Else
             selta1012_qr(0) = Msk_caseno.FormattedText
             selta1012_rs.MoreResults
             selta1012_rs.Requery
        End If
        If selta1012_rs.RowCount > 0 Then
            UPDTA1012_QR(0) = Trim(txt_indctinvstg.Text)
            UPDTA1012_QR(1) = Msk_caseno.FormattedText
            UPDTA1012_QR.Execute
        Else
            insta1012_qr(0) = Msk_caseno.FormattedText
            insta1012_qr(1) = Trim(txt_indctinvstg.Text)
            insta1012_qr.Execute
        End If
        End If
        '*** UPDATING THE PROVISIONAL DIAGNOSIS
        '*** updating the clinical history
        If SELCLINHISTFLAG = False Then
            SELCLINHIST_QR(0) = Msk_caseno.FormattedText
            Set SELCLINHIST_RS = SELCLINHIST_QR.OpenResultset(1, 3)
            SELCLINHISTFLAG = True
        Else
            SELCLINHIST_QR(0) = Msk_caseno.FormattedText
            SELCLINHIST_RS.MoreResults
            SELCLINHIST_RS.Requery
        End If
        'If SELCLINHIST_RS.RowCount > 0 Then
'            updclinhist_qr(0) = Trim(txt_clinhist.Text)
'            updclinhist_qr(1) = Trim(Msk_caseno.FormattedText)
'            updclinhist_qr.Execute
'        Else
            insclinhist_qr(0) = Trim(Msk_caseno.FormattedText)
            insclinhist_qr(1) = GET_REVDATE(Trim(DATE11))
            insclinhist_qr(2) = Trim(time1)
            If Trim(txt_clinhist.Text) = "NYD" Then
                insclinhist_qr(3) = "NYD"
            Else
                insclinhist_qr(3) = Trim(txt_clinhist.Text)
            End If
            insclinhist_qr.Execute
        'End If
        '*** UPDATING THE RREQUISITN TABLE
        updrrequisitn_qr(0) = UCase(Trim(cmb_refdoc.Text))
        updrrequisitn_qr(1) = req_chglen(cmb_reqno.Text)
        updrrequisitn_qr.Execute
        If SELRREQUISITNFLAG2 = False Then
                SELRREQUISITN_QR2(0) = req_chglen(Trim(cmb_reqno.Text))
                Set SELRREQUISITN_RS2 = SELRREQUISITN_QR2.OpenResultset(1, 3)
                SELRREQUISITNFLAG2 = True
           Else
                SELRREQUISITN_QR2(0) = req_chglen(Trim(cmb_reqno.Text))
                SELRREQUISITN_RS2.MoreResults
                SELRREQUISITN_RS2.Requery
           End If
        If SELRREQUISITN_RS2.RowCount > 0 Then
                If IsNull(SELRREQUISITN_RS2("reqnstat")) Then
        
                If CMB_CGNIP.Text = "VPP" Or CMB_CGNIP.Text = "LR" Or CMB_CGNIP.Text = "BAR" Or CMB_CGNIP.Text = "DrOp" Then
                    If CMB_CGNIP.Text = "VPP" Then
                        UPDRREQQR(0) = Null
                        UPDRREQQR(1) = Null
                        UPDRREQQR(2) = Null
                        UPDRREQQR(3) = 1
                        UPDRREQQR(4) = req_chglen(cmb_reqno.Text)
                        UPDRREQQR.Execute
                    End If
                    If CMB_CGNIP.Text = "LR" Then
                        UPDRREQQR(0) = "LR"
                        UPDRREQQR(1) = GET_REVDATE(Trim(DATE11))
                        UPDRREQQR(2) = "P"
                        UPDRREQQR(3) = 2
                        UPDRREQQR(4) = req_chglen(cmb_reqno.Text)
                        UPDRREQQR.Execute
                    End If
                    If CMB_CGNIP.Text = "BAR" Then
                        UPDRREQQR(0) = "BAR"
                        UPDRREQQR(1) = GET_REVDATE(Trim(DATE11))
                        UPDRREQQR(2) = "P"
                        UPDRREQQR(3) = 3
                        UPDRREQQR(4) = req_chglen(cmb_reqno.Text)
                        UPDRREQQR.Execute
                    End If
                    If CMB_CGNIP.Text = "DrOp" Then
                        UPDRREQQR(0) = "DrOp"
                        UPDRREQQR(1) = GET_REVDATE(Trim(DATE11))
                        UPDRREQQR(2) = "P"
                        UPDRREQQR(3) = 4
                        UPDRREQQR(4) = req_chglen(cmb_reqno.Text)
                        UPDRREQQR.Execute
                    End If
                End If
              End If
       End If
        
        
        
        '*** UPDATING THE RREQUISITN TABLE
End Sub

Public Sub CGENABL(str As String)
If str = "ENA" Then
    CMB_CGPROCEDURE.Enabled = True
    cbo_proccode.Enabled = True
    cmb_cnsite.Enabled = True
    cmb_claterality.Enabled = True
    txt_trtdetail.Enabled = True
    txt_cnatmeterial.Enabled = True
    cmb_cnoslides.Enabled = True
    cmb_cyoutspecimen.Enabled = True
    cmb_coutslides.Enabled = True
    txt_scopyfind.Enabled = False
    MSK_LMP.Enabled = True
    cmb_cpreinv.Enabled = True
      
End If
If str = "DIS" Then
    CMB_CGPROCEDURE.Enabled = False
    cbo_proccode.Enabled = False
    cmb_cnsite.Enabled = False
    cmb_claterality.Enabled = False
    txt_trtdetail.Enabled = False
    txt_cnatmeterial.Enabled = False
    cmb_cnoslides.Enabled = False
    cmb_cyoutspecimen.Enabled = False
    cmb_coutslides.Enabled = False
    txt_scopyfind.Enabled = False
    MSK_LMP.Enabled = False
    cmb_cpreinv.Enabled = True
      
End If
If str = "MOD" Then
    CMB_CGPROCEDURE.Enabled = False
    cbo_proccode.Enabled = False
    cmb_cnsite.Enabled = True
    cmb_claterality.Enabled = True
    txt_trtdetail.Enabled = True
    txt_cnatmeterial.Enabled = True
    cmb_cnoslides.Enabled = True
    cmb_cyoutspecimen.Enabled = True
    cmb_coutslides.Enabled = True
    txt_scopyfind.Enabled = False
    MSK_LMP.Enabled = True
    cmb_cpreinv.Enabled = True
    cmb_refdoc.Enabled = True
    txt_indctinvstg.Enabled = True
End If
End Sub

Private Sub txt_scopyfind_GotFocus()
On Error GoTo m1:
WSCOPYCY.Visible = False
Set DSCOPYCY = WSCOPYCY.ActiveDocument
DSCOPYCY.Content = ""
DSCOPYCY.Content = txt_scopyfind
WSCOPYCY.Caption = "Scopy Finding"
WSCOPYCY.WindowState = wdWindowStateMaximize
DSCOPYCY.Content = txt_scopyfind
WSCOPYCY.Visible = True
WSCOPYCY.Activate
m1:
On Error Resume Next
Set WSCOPYCY = CreateObject("Word.Application")
WSCOPYCY.Documents.add ("C:\DISREPORTSEC\REPORT.DOT")
WSCOPYCY.Caption = "Scopy Finding"
WSCOPYCY.WindowState = wdWindowStateMaximize
Set DSCOPYCY = WSCOPYCY.ActiveDocument
DSCOPYCY.Content = Trim(txt_scopyfind.Text)
WSCOPYCY.Visible = True
WSCOPYCY.Activate
End Sub

Private Sub txt_scopyfind_LostFocus()
On Error GoTo m2:
Set DSCOPYCY = WSCOPYCY.ActiveDocument
txt_scopyfind.Text = DSCOPYCY.Content
If Len(Trim(txt_scopyfind.Text)) >= 500 Then
    MsgBox "Scopyfinding Max. of 500 Characters"
    txt_scopyfind.SetFocus
    Exit Sub
End If
'flagimp = False
DSCOPYCY.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'DSCOPYCY.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
DSCOPYCY.Close (wdDoNotSaveChanges)
WSCOPYCY.quit
'W1.Visible = False
Exit Sub
m2:
MsgBox Err.Description
End Sub

Private Sub txt_trtdetail_GotFocus()
On Error GoTo m1:
WTREATCY.Visible = False
Set DTREATCY = WTREATCY.ActiveDocument
DTREATCY.Content = ""
DTREATCY.Content = txt_trtdetail
WTREATCY.Caption = "Treament Details"
WTREATCY.WindowState = wdWindowStateMaximize
DTREATCY.Content = txt_trtdetail
WTREATCY.Visible = True
WTREATCY.Activate
m1:
On Error Resume Next
Set WTREATCY = CreateObject("Word.Application")
WTREATCY.Documents.add ("C:\DISREPORTSEC\REPORT.DOT")
WTREATCY.Caption = "Treatment Details"
WTREATCY.WindowState = wdWindowStateMaximize
Set DTREATCY = WTREATCY.ActiveDocument
DTREATCY.Content = Trim(txt_trtdetail.Text)
WTREATCY.Visible = True
WTREATCY.Activate
End Sub

Private Sub txt_trtdetail_LostFocus()
On Error GoTo m2:
Set DTREATCY = WTREATCY.ActiveDocument
txt_trtdetail.Text = DTREATCY.Content
If Len(Trim(txt_trtdetail.Text)) >= 1500 Then
    MsgBox "Treatment Details Max. of 1500 Characters"
    txt_trtdetail.SetFocus
    Exit Sub
End If
'flagimp = False
DTREATCY.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'DTREATCY.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
DTREATCY.Close (wdDoNotSaveChanges)
WTREATCY.quit
'W1.Visible = False
Exit Sub
m2:
MsgBox Err.Description
End Sub

Public Sub LRBR()

End Sub

Public Sub BC()
surgicalPathology.Visible = False
immunohistochemistry.Visible = False
frm_hematopathology.Visible = False
BIOCHEMISTRY1.Visible = False
bceltm.Visible = True
 Label10.Visible = False
 Label11.Visible = False
 cmb_cyprocedure.Visible = False
 cmb_cyprocode.Visible = False
 
'cardiology.Visible = False
Microbiology.Visible = False
fnac.Visible = False
CYTOLOGY.Visible = FASLE
BIOCHEMISTRY1.show
Em_lab.Hide
TUMOURMARKER.Hide
BIOCHEMISTRY1.Msk_caseno.Text = Msk_caseno.FormattedText
BIOCHEMISTRY1.txt_name.Text = txt_name.Text

lbl_req.Left = 3500
lbl_req.Caption = "REQUISITION - BIOCHEMISTRY"
pathology.Caption = "Pathology - Biochemistry"


        
End Sub

Public Sub EL()
        surgicalPathology.Visible = False
        immunohistochemistry.Visible = False
        frm_hematopathology.Visible = False
        biochemistry.Visible = False
        bceltm.Visible = True
        Label10.Visible = False
        Label11.Visible = False
        cmb_cyprocedure.Visible = False
         cmb_cyprocode.Visible = False
'        cardiology.Visible = False
        Microbiology.Visible = False
        fnac.Visible = False
        CYTOLOGY.Visible = FASLE
        biochemistry.Hide
'        If lbl_icureq.Visible = True Then
'            Em_lab.Frame10.Visible = True
'        Else
'            Em_lab.Frame10.Visible = False
'        End If
        Em_lab.show
        TUMOURMARKER.Hide
        hemcbc.Hide
        HEMBF.Hide
        HEMBMA.Hide
        CMB_TEXT.clear
        Em_lab.MaskEdBox1.Text = Msk_caseno.FormattedText
        Em_lab.txt_name.Text = txt_name.Text
        lbl_req.Caption = "REQUISITION - EMERGENCY LAB"
        pathology.Caption = "Pathology - Emergency Lab"
End Sub

Public Sub TM()
        surgicalPathology.Visible = False
        immunohistochemistry.Visible = False
        frm_hematopathology.Visible = False
        biochemistry.Visible = False
        bceltm.Visible = True
        Label10.Visible = False
        Label11.Visible = False
        cmb_cyprocedure.Visible = False
        cmb_cyprocode.Visible = False
'        cardiology.Visible = False
        Microbiology.Visible = False
        fnac.Visible = False
        CYTOLOGY.Visible = FASLE
        biochemistry.Hide
        Em_lab.Hide
        TUMOURMARKER.show
        TUMOURMARKER.Msk_caseno.Text = Msk_caseno.FormattedText
        TUMOURMARKER.txt_name.Text = txt_name.Text
        lbl_req.Caption = "REQUISITION - TUMOUR MARKER"
        pathology.Caption = "Pathology - Tumour Marker"
        TUMOURMARKER.txt_clininfo.Text = txt_clinhist.Text
End Sub

Public Sub BCASSIGN()
Dim durill As String
Dim I As Integer

On Error GoTo MSG:

'*** INSERTING INT0 RREQUISITION
INSRREQUISITN_QR(0) = req_chglen(Trim(txt_bcreqno.Text))
INSRREQUISITN_QR(1) = GET_REVDATE(CStr(DATE11))
INSRREQUISITN_QR(2) = Trim(txt_time.Text)
INSRREQUISITN_QR(3) = Trim(Msk_caseno.FormattedText)
INSRREQUISITN_QR(4) = Trim(txt_dept.Text)
INSRREQUISITN_QR(5) = Trim(txt_modcode.Text)
INSRREQUISITN_QR(6) = Trim(txt_status.Text)
INSRREQUISITN_QR(7) = Trim(txt_catgory.Text)
INSRREQUISITN_QR(8) = "8888"
INSRREQUISITN_QR(9) = UCase(Trim(cmb_refdoc.Text))
'INSRREQUISITN_QR(9) = Trim(txt_pridocser.Text)
  
'Added by Milind on 27-03-2006 for inserting date for the samples collected at outside
    If Check1.Value = 1 Then
        Dim Dt As String
        Dt = "18/02/1980"
        INSRREQUISITN_QR(15) = Format(CDate(Dt), "dd/mm/yyyy")
    Else
        INSRREQUISITN_QR(15) = Null
    End If
'upto here

    If SELRREQUISITNFLAG = False Then
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
        SELRREQUISITNFLAG = True
    Else
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        SELRREQUISITN_RS.MoreResults
        SELRREQUISITN_RS.Requery
    End If
    If SELRREQUISITN_RS.RowCount > 0 Then
        INSRREQUISITN_QR(10) = "YES"
    Else
        INSRREQUISITN_QR(10) = "NO"
    End If
If opt_patient.Value = True Then
    NO_CHARGE
    ''' Added by Sandhya dated 26.6.2009
    ''' for NC category & charged patient
    If txt_catgory.Text = "NC" And No_Chargeflag = True Then
        INSRREQUISITN_QR(11) = "NOCHARGE"
        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
        INSRREQUISITN_QR(13) = "P"
        'No_Chargeflag = False
    ''ICU
    ElseIf lbl_icureq.Visible = True Then
        'If Trim(CMB_ICU1.Text) = "YES" Then
         INSRREQUISITN_QR(11) = Null
         INSRREQUISITN_QR(12) = Null
        ' INSRREQUISITN_QR(13) = "P"
        ' INSRREQUISITN_QR(13) = "E"
       ' INSRREQUISITN_QR(13) = " "
        INSRREQUISITN_QR(13) = Null
        'End If
    ''ICU
    Else
        INSRREQUISITN_QR(11) = Null
        INSRREQUISITN_QR(12) = Null
        INSRREQUISITN_QR(13) = Null
    End If
End If
If opt_staff.Value = True Then
        INSRREQUISITN_QR(11) = "NOCHARGE"
        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
        INSRREQUISITN_QR(13) = "P"
End If
INSRREQUISITN_QR(14) = 0

INSRREQUISITN_QR.Execute
    '*** INSERTING INT0 RREQUISITION
    
    '*** INSERTING INTO DMULPROCE TABLE
    If CMB_TEXT.ListCount <> 0 Then
        For I = 0 To CMB_TEXT.ListCount - 1
            INSDMULPROCE_QR(0) = req_chglen(txt_bcreqno.Text)
            INSDMULPROCE_QR(1) = CMB_TEXT.List(I)
            INSDMULPROCE_QR(2) = GET_REVDATE(DATE11)
            INSDMULPROCE_QR(3) = Msk_caseno.FormattedText
            INSDMULPROCE_QR.Execute
        Next I
    End If
    '*** INSERTING INTO DMULPROCE TABLE
    
    '**** INSERTING INTO DREQUISITION

INSDREQUISITN_QR(0) = req_chglen(Trim(txt_bcreqno.Text))
INSDREQUISITN_QR(1) = " "
INSDREQUISITN_QR(2) = Trim(Msk_caseno.FormattedText)
INSDREQUISITN_QR(3) = Null

INSDREQUISITN_QR(4) = Null
dur_ill = txt_year & "/" & Trim(CMB_DURILLMONTH.Text)
INSDREQUISITN_QR(5) = dur_ill
INSDREQUISITN_QR(6) = bodyf
INSDREQUISITN_QR(7) = Null
INSDREQUISITN_QR(8) = Null
INSDREQUISITN_QR(9) = Null
INSDREQUISITN_QR(10) = Null
INSDREQUISITN_QR(11) = Null
INSDREQUISITN_QR(12) = Null
INSDREQUISITN_QR(13) = Null
INSDREQUISITN_QR(14) = Trim(cmb_refdoc.Text)
INSDREQUISITN_QR(15) = Trim(cmb_reftodoc.Text)
INSDREQUISITN_QR(16) = Null
If Trim(txt_status.Text) = "IP" Then
    INSDREQUISITN_QR(17) = Trim(txt_ward.Text)
    INSDREQUISITN_QR(18) = Trim(txt_bedno.Text)
Else
    INSDREQUISITN_QR(17) = Null
    INSDREQUISITN_QR(18) = Null
End If
INSDREQUISITN_QR(19) = " "
If Len(Trim(meterial)) <> 0 Then

    INSDREQUISITN_QR(20) = Null
    INSDREQUISITN_QR(20) = meterial
    
Else
    INSDREQUISITN_QR(20) = Null
    INSDREQUISITN_QR(20) = ""
End If


INSDREQUISITN_QR(21) = Null
INSDREQUISITN_QR.Execute
'**** INSERTING INTO DREQUISITION

 '*** UPDATING THE PROVISIONAL DIAGNOSIS
 
If opt_patient.Value = True Then
If SELTA1012FLAG = False Then
    selta1012_qr(0) = Msk_caseno.FormattedText
    Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
    SELTA1012FLAG = True
Else
     selta1012_qr(0) = Msk_caseno.FormattedText
     selta1012_rs.MoreResults
     selta1012_rs.Requery
End If
If selta1012_rs.RowCount > 0 Then
    UPDTA1012_QR(0) = Trim(txt_indctinvstg.Text)
    UPDTA1012_QR(1) = Msk_caseno.FormattedText
    UPDTA1012_QR.Execute
Else
    insta1012_qr(0) = Msk_caseno.FormattedText
    insta1012_qr(1) = Trim(txt_indctinvstg.Text)
    insta1012_qr.Execute
End If
End If
'*** UPDATING THE PROVISIONAL DIAGNOSIS
'*** INSERTING INTO DMULPROCE FOR COLLECTION AND DEPT COLLECTION SCREEN
''    INSDMULPROCE_QR(0) = req_chglen(txt_spreqno.Text)
''    INSDMULPROCE_QR(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
''    INSDMULPROCE_QR(2) = Msk_caseno.FormattedText
''    INSDMULPROCE_QR(3) = cmb_procedurecode.Text
''    INSDMULPROCE_QR.Execute

'***INSERTING INTO DMULPROCE FOR COLLECTION AND DEPT COLLECTION SCREEN
'***INSERTING INTO DREQSTAFFD TABLE
If opt_staff.Value = True Then
    INSDREQSTAFFD_QR(0) = req_chglen(txt_bcreqno.Text)
    INSDREQSTAFFD_QR(1) = Msk_caseno.Text
    INSDREQSTAFFD_QR(2) = Trim(txt_name.Text)
    INSDREQSTAFFD_QR(3) = Trim(txt_Age)
    INSDREQSTAFFD_QR(4) = Trim(txt_sex)
    If Len(Trim(cmb_dependents.Text)) = 0 Then
        INSDREQSTAFFD_QR(5) = "SELF"
    Else
        I = InStr(cmb_dependents, "$")
        INSDREQSTAFFD_QR(5) = Mid(cmb_dependents, I + 1, 10)
    End If
    INSDREQSTAFFD_QR.Execute
End If
'***INSERTING INTO DREQSTAFFD TABLE
    If SELCLINHISTFLAG = False Then
        SELCLINHIST_QR(0) = Msk_caseno.FormattedText
        Set SELCLINHIST_RS = SELCLINHIST_QR.OpenResultset(1, 3)
        SELCLINHISTFLAG = True
    Else
        SELCLINHIST_QR(0) = Msk_caseno.FormattedText
        SELCLINHIST_RS.MoreResults
        SELCLINHIST_RS.Requery
    End If
'    If SELCLINHIST_RS.RowCount > 0 Then
'        updclinhist_qr(0) = Trim(txt_clinhist.Text)
'        updclinhist_qr(1) = Trim(Msk_caseno.Text)
'        updclinhist_qr.Execute
'    Else
        insclinhist_qr(0) = Trim(Msk_caseno.FormattedText)
        insclinhist_qr(1) = GET_REVDATE(Trim(DATE11))
        insclinhist_qr(2) = Trim(time1)
        If Trim(txt_clinhist.Text) = "NYD" Then
            insclinhist_qr(3) = "NYD"
        Else
            insclinhist_qr(3) = Trim(txt_clinhist.Text)
        End If
        insclinhist_qr.Execute
    'End If
''UPDATE DREQUISITN TABLE
If SELDISICU_FLAG = False Then
    SELDISICU_QR(0) = Trim(ipadd)
    Set SELDISICU_RS = SELDISICU_QR.OpenResultset(1, 3)
    SELDISICU_FLAG = True
Else
    SELDISICU_QR(0) = Trim(ipadd)
    SELDISICU_RS.MoreResults
    SELDISICU_RS.Requery
End If

If SELDISICU_RS.RowCount > 0 Then
    If Trim(SELDISICU_RS(1)) = "10.3.1.56" Then
        UPDICU_QR(0) = "YR"
        UPDICU_QR(1) = req_chglen(Trim(txt_bcreqno.Text))
        UPDICU_QR.Execute
    Else
        UPDICU_QR(0) = "YES"
        UPDICU_QR(1) = req_chglen(Trim(txt_bcreqno.Text))
        UPDICU_QR.Execute
    End If
End If

    
''UPDATE DREQUISITN TABLE

Data1.UpdateRecord
Data1.Recordset.Bookmark = Data1.Recordset.LastModified
Exit Sub
MSG:
MsgBox Err.Description

'CSF_FLAG = False
End Sub

Public Sub ELASSIGN()

Dim durill As String
Dim I As Integer

On Error GoTo MSG:
'*** INSERTING INT0 RREQUISITION
INSRREQUISITN_QR(0) = req_chglen(Trim(txt_bcreqno.Text))

INSRREQUISITN_QR(1) = GET_REVDATE(CStr(DATE11))
INSRREQUISITN_QR(2) = Trim(txt_time.Text)
INSRREQUISITN_QR(3) = Trim(Msk_caseno.FormattedText)
INSRREQUISITN_QR(4) = Trim(txt_dept.Text)
INSRREQUISITN_QR(5) = Trim(txt_modcode.Text)
INSRREQUISITN_QR(6) = Trim(txt_status.Text)
INSRREQUISITN_QR(7) = Trim(txt_catgory.Text)
INSRREQUISITN_QR(8) = "8888"
INSRREQUISITN_QR(9) = UCase(Trim(cmb_refdoc.Text))
'INSRREQUISITN_QR(9) = Trim(txt_pridocser.Text)

'Added by Milind on 27-03-2006 for inserting date for the samples collected at outside
    If Check1.Value = 1 Then
        Dim Dt As String
        Dt = "18/02/1980"
        INSRREQUISITN_QR(15) = Format(CDate(Dt), "dd/mm/yyyy")
    Else
        INSRREQUISITN_QR(15) = Null
    End If
'upto here

    If SELRREQUISITNFLAG = False Then
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
        SELRREQUISITNFLAG = True
    Else
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        SELRREQUISITN_RS.MoreResults
        SELRREQUISITN_RS.Requery
    End If
    If SELRREQUISITN_RS.RowCount > 0 Then
        INSRREQUISITN_QR(10) = "YES"
    Else
        INSRREQUISITN_QR(10) = "NO"
    End If
If opt_patient.Value = True Then
    'ADDED BY MILIND ON 23-08-06
    If CMB_TEXT.ListCount <> 0 Then
        For I = 0 To CMB_TEXT.ListCount - 1
            'MILIND FOR CHECKING PREVENTIVE ONCOLOGY FREE CODES
            If CMB_TEXT.List(I) = "F904" Or CMB_TEXT.List(I) = "F905" Or CMB_TEXT.List(I) = "F911" Then
                NOCHARGEEL
            End If
            ''' Added by Sandhya dated 26.6.2009
            ''' for NC category & charged patient
             NOCHARGEEL
            If txt_catgory.Text = "NC" Or No_Chargeflag = True Then
                        INSRREQUISITN_QR(11) = "NOCHARGE"
                        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
                        INSRREQUISITN_QR(13) = "P"
                        'NO_CHARGE
            Else
                    INSRREQUISITN_QR(11) = Null
                    INSRREQUISITN_QR(12) = Null
                    INSRREQUISITN_QR(13) = Null
            End If
          
        Next I
'        If No_Chargeflag = True Then
'            INSRREQUISITN_QR(11) = "NOCHARGE"
'            INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
'            INSRREQUISITN_QR(13) = "P"
'            No_Chargeflag = False
'        Else
'            INSRREQUISITN_QR(11) = Null
'            INSRREQUISITN_QR(12) = Null
'            INSRREQUISITN_QR(13) = Null
'        End If
    End If
'    If txt_catgory.Text = "NC" Then
'        INSRREQUISITN_QR(11) = "NOCHARGE"
'        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
'        INSRREQUISITN_QR(13) = "P"
'        'No_Chargeflag = False
'    ''ICU
'    ElseIf lbl_icureq.Visible = True Then
'        'If Trim(CMB_ICU1.Text) = "YES" Then
'            INSRREQUISITN_QR(11) = Null
'            INSRREQUISITN_QR(12) = Null
'           ' INSRREQUISITN_QR(13) = "P"
'           ' INSRREQUISITN_QR(13) = "E"
'           'INSRREQUISITN_QR(13) = " "
'            INSRREQUISITN_QR(13) = Null
'        'End If
'    Else
'        INSRREQUISITN_QR(11) = Null
'        INSRREQUISITN_QR(12) = Null
'        INSRREQUISITN_QR(13) = Null
'    End If
End If
If opt_staff.Value = True Then
        INSRREQUISITN_QR(11) = "NOCHARGE"
        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
        INSRREQUISITN_QR(13) = "P"
End If
INSRREQUISITN_QR(14) = 0

INSRREQUISITN_QR.Execute
    '*** INSERTING INT0 RREQUISITION
    '*** INSERTING INTO DMULPROCE TABLE
    If CMB_TEXT.ListCount <> 0 Then
        For I = 0 To CMB_TEXT.ListCount - 1
            INSDMULPROCE_QR(0) = req_chglen(txt_bcreqno.Text)
            INSDMULPROCE_QR(1) = CMB_TEXT.List(I)
            INSDMULPROCE_QR(2) = GET_REVDATE(DATE11)
            INSDMULPROCE_QR(3) = Msk_caseno.FormattedText
            INSDMULPROCE_QR.Execute
        Next I
    End If
    '*** INSERTING INTO DMULPROCE TABLE
    '**** INSERTING INTO DREQUISITION

INSDREQUISITN_QR(0) = req_chglen(Trim(txt_bcreqno.Text))
INSDREQUISITN_QR(1) = " "
INSDREQUISITN_QR(2) = Trim(Msk_caseno.FormattedText)
INSDREQUISITN_QR(3) = Null

INSDREQUISITN_QR(4) = Null
dur_ill = txt_year & "/" & Trim(CMB_DURILLMONTH.Text)
INSDREQUISITN_QR(5) = dur_ill
INSDREQUISITN_QR(6) = Null
INSDREQUISITN_QR(7) = Null
INSDREQUISITN_QR(8) = Null
INSDREQUISITN_QR(9) = Null
INSDREQUISITN_QR(10) = Null
INSDREQUISITN_QR(11) = Null
INSDREQUISITN_QR(12) = Null
INSDREQUISITN_QR(13) = Null
INSDREQUISITN_QR(14) = Trim(cmb_refdoc.Text)
INSDREQUISITN_QR(15) = Trim(cmb_reftodoc.Text)
INSDREQUISITN_QR(16) = Null
If Trim(txt_status.Text) = "IP" Then
    INSDREQUISITN_QR(17) = Trim(txt_ward.Text)
    INSDREQUISITN_QR(18) = Trim(txt_bedno.Text)
Else
    INSDREQUISITN_QR(17) = Null
    INSDREQUISITN_QR(18) = Null
End If
INSDREQUISITN_QR(19) = " "
If Len(Trim(meterial)) > 0 Then
    INSDREQUISITN_QR(20) = Null
    
    INSDREQUISITN_QR(20) = meterial
Else
    INSDREQUISITN_QR(20) = Null
    INSDREQUISITN_QR(20) = ""
End If
INSDREQUISITN_QR(21) = Null
INSDREQUISITN_QR.Execute

'**** INSERTING INTO DREQUISITION

'*** UPDATING THE PROVISIONAL DIAGNOSIS
 
If opt_patient.Value = True Then
    If SELTA1012FLAG = False Then
        selta1012_qr(0) = Msk_caseno.FormattedText
        Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
        SELTA1012FLAG = True
    Else
        selta1012_qr(0) = Msk_caseno.FormattedText
        selta1012_rs.MoreResults
        selta1012_rs.Requery
    End If
    If selta1012_rs.RowCount > 0 Then
        UPDTA1012_QR(0) = Trim(txt_indctinvstg.Text)
        UPDTA1012_QR(1) = Msk_caseno.FormattedText
        UPDTA1012_QR.Execute
    Else
        insta1012_qr(0) = Msk_caseno.FormattedText
        insta1012_qr(1) = Trim(txt_indctinvstg.Text)
        insta1012_qr.Execute
    End If
End If
'*** UPDATING THE PROVISIONAL DIAGNOSIS

'***INSERTING INTO DREQSTAFFD TABLE
If opt_staff.Value = True Then
    INSDREQSTAFFD_QR(0) = req_chglen(txt_bcreqno.Text)
    INSDREQSTAFFD_QR(1) = Msk_caseno.Text
    INSDREQSTAFFD_QR(2) = Trim(txt_name.Text)
    INSDREQSTAFFD_QR(3) = Trim(txt_Age)
    INSDREQSTAFFD_QR(4) = Trim(txt_sex)
    If Len(Trim(cmb_dependents.Text)) = 0 Then
        INSDREQSTAFFD_QR(5) = "SELF"
    Else
        I = InStr(cmb_dependents, "$")
        INSDREQSTAFFD_QR(5) = Mid(cmb_dependents, I + 1, 10)
    End If
    INSDREQSTAFFD_QR.Execute
End If
'***INSERTING INTO DREQSTAFFD TABLE
    If SELCLINHISTFLAG = False Then
        SELCLINHIST_QR(0) = Msk_caseno.FormattedText
        Set SELCLINHIST_RS = SELCLINHIST_QR.OpenResultset(1, 3)
        SELCLINHISTFLAG = True
    Else
        SELCLINHIST_QR(0) = Msk_caseno.FormattedText
        SELCLINHIST_RS.MoreResults
        SELCLINHIST_RS.Requery
    End If
'    If SELCLINHIST_RS.RowCount > 0 Then
'        updclinhist_qr(0) = Trim(txt_clinhist.Text)
'        updclinhist_qr(1) = Trim(Msk_caseno.Text)
'        updclinhist_qr.Execute
'    Else
        insclinhist_qr(0) = Trim(Msk_caseno.FormattedText)
        insclinhist_qr(1) = GET_REVDATE(Trim(DATE11))
        insclinhist_qr(2) = Trim(time1)
        If Trim(txt_clinhist.Text) = "NYD" Then
            insclinhist_qr(3) = "NYD"
        Else
            insclinhist_qr(3) = Trim(txt_clinhist.Text)
        End If
        insclinhist_qr.Execute
    'End If

    ''UPDATE DREQUISITN TABLE
    If SELDISICU_FLAG = False Then
        SELDISICU_QR(0) = Trim(ipadd)
        Set SELDISICU_RS = SELDISICU_QR.OpenResultset(1, 3)
        SELDISICU_FLAG = True
    Else
        SELDISICU_QR(0) = Trim(ipadd)
        SELDISICU_RS.MoreResults
        SELDISICU_RS.Requery
    End If
    
    If SELDISICU_RS.RowCount > 0 Then
        If Trim(SELDISICU_RS(1)) = "10.3.1.56" Then
            UPDICU_QR(0) = "YR"
            UPDICU_QR(1) = req_chglen(Trim(txt_bcreqno.Text))
            UPDICU_QR.Execute
        Else
            UPDICU_QR(0) = "YES"
            UPDICU_QR(1) = req_chglen(Trim(txt_bcreqno.Text))
            UPDICU_QR.Execute
        End If
    End If
    
'    If Not SELDISICU_RS.EOF Then
'        UPDICU_QR(0) = "YES"
'        UPDICU_QR(1) = req_chglen(Trim(txt_bcreqno.Text))
'        UPDICU_QR.Execute
'    End If
    
''UPDATE DREQUISITN TABLE
Data1.UpdateRecord
Data1.Recordset.Bookmark = Data1.Recordset.LastModified
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Public Sub TMASSIGN()
Dim durill As String
Dim I As Integer

On Error GoTo MSG:
'*** INSERTING INT0 RREQUISITION
INSRREQUISITN_QR(0) = req_chglen(Trim(txt_bcreqno.Text))
INSRREQUISITN_QR(1) = GET_REVDATE(CStr(DATE11))
INSRREQUISITN_QR(2) = Trim(txt_time.Text)
INSRREQUISITN_QR(3) = Trim(Msk_caseno.FormattedText)
INSRREQUISITN_QR(4) = Trim(txt_dept.Text)
INSRREQUISITN_QR(5) = Trim(txt_modcode.Text)
INSRREQUISITN_QR(6) = Trim(txt_status.Text)
INSRREQUISITN_QR(7) = Trim(txt_catgory.Text)
INSRREQUISITN_QR(8) = "8888"
INSRREQUISITN_QR(9) = UCase(Trim(cmb_refdoc.Text))
'INSRREQUISITN_QR(9) = Trim(txt_pridocser.Text)

'Added by Milind on 27-03-2006 for inserting date for the samples collected at outside
    If Check1.Value = 1 Then
        Dim Dt As String
        Dt = "18/02/1980"
        INSRREQUISITN_QR(15) = Format(CDate(Dt), "dd/mm/yyyy")
    Else
        INSRREQUISITN_QR(15) = Null
    End If
'upto here

    If SELRREQUISITNFLAG = False Then
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
        SELRREQUISITNFLAG = True
    Else
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        SELRREQUISITN_RS.MoreResults
        SELRREQUISITN_RS.Requery
    End If
    If SELRREQUISITN_RS.RowCount > 0 Then
        INSRREQUISITN_QR(10) = "YES"
    Else
        INSRREQUISITN_QR(10) = "NO"
    End If
If opt_patient.Value = True Then
    'ADDED BY MILIND ON 23-08-06
    If CMB_TEXT.ListCount <> 0 Then
        For I = 0 To CMB_TEXT.ListCount - 1
            'MILIND FOR CHECKING PREVENTIVE ONCOLOGY FREE CODES
            If CMB_TEXT.List(I) = "F820" Then
                NOCHARGETM
            End If
            ''' Added by Sandhya dated 26.6.2009
            ''' for NC category & charged patient
            NOCHARGETM
            If txt_catgory.Text = "NC" And No_Chargeflag = True Then
                        INSRREQUISITN_QR(11) = "NOCHARGE"
                        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
                        INSRREQUISITN_QR(13) = "P"
                        'NO_CHARGE
            Else
                    INSRREQUISITN_QR(11) = Null
                    INSRREQUISITN_QR(12) = Null
                    INSRREQUISITN_QR(13) = Null
            End If
        Next I
    End If
'    'NO_CHARGE
'    If txt_catgory.Text = "NC" Then
'        INSRREQUISITN_QR(11) = "NOCHARGE"
'        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
'        INSRREQUISITN_QR(13) = "P"
'        'No_Chargeflag = False
'    ''ICU
'    ElseIf lbl_icureq.Visible = True Then
'        'If Trim(CMB_ICU1.Text) = "YES" Then
'            INSRREQUISITN_QR(11) = Null
'            INSRREQUISITN_QR(12) = Null
'            'INSRREQUISITN_QR(13) = "P"
'            'INSRREQUISITN_QR(13) = " "
'             INSRREQUISITN_QR(13) = Null
'
'        'End If
'    Else
'
'            INSRREQUISITN_QR(11) = Null
'            INSRREQUISITN_QR(12) = Null
'            INSRREQUISITN_QR(13) = Null
'
'    End If
End If
If opt_staff.Value = True Then
        INSRREQUISITN_QR(11) = "NOCHARGE"
        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
        INSRREQUISITN_QR(13) = "P"
End If
INSRREQUISITN_QR(14) = 0

INSRREQUISITN_QR.Execute
    '*** INSERTING INT0 RREQUISITION
    '*** INSERTING INTO DMULPROCE TABLE
    If CMB_TEXT.ListCount <> 0 Then
        For I = 0 To CMB_TEXT.ListCount - 1
            INSDMULPROCE_QR(0) = req_chglen(txt_bcreqno.Text)
            INSDMULPROCE_QR(1) = CMB_TEXT.List(I)
            INSDMULPROCE_QR(2) = GET_REVDATE(DATE11)
            INSDMULPROCE_QR(3) = Msk_caseno.FormattedText
            INSDMULPROCE_QR.Execute
        Next I
    End If
    '*** INSERTING INTO DMULPROCE TABLE
    '**** INSERTING INTO DREQUISITION

INSDREQUISITN_QR(0) = req_chglen(Trim(txt_bcreqno.Text))
INSDREQUISITN_QR(1) = " "
INSDREQUISITN_QR(2) = Trim(Msk_caseno.FormattedText)
INSDREQUISITN_QR(3) = Null
INSDREQUISITN_QR(4) = Null
dur_ill = txt_year & "/" & Trim(CMB_DURILLMONTH.Text)
INSDREQUISITN_QR(5) = dur_ill
INSDREQUISITN_QR(6) = AM
INSDREQUISITN_QR(7) = Null
INSDREQUISITN_QR(8) = Null
INSDREQUISITN_QR(9) = Null
INSDREQUISITN_QR(10) = Null
INSDREQUISITN_QR(11) = Null
INSDREQUISITN_QR(12) = Null
INSDREQUISITN_QR(13) = Null
INSDREQUISITN_QR(14) = Trim(cmb_refdoc.Text)
INSDREQUISITN_QR(15) = Trim(cmb_reftodoc.Text)
INSDREQUISITN_QR(16) = Null
If Trim(txt_status.Text) = "IP" Then
    INSDREQUISITN_QR(17) = Trim(txt_ward.Text)
    INSDREQUISITN_QR(18) = Trim(txt_bedno.Text)
Else
    INSDREQUISITN_QR(17) = Null
    INSDREQUISITN_QR(18) = Null
End If

INSDREQUISITN_QR(19) = DDG

If Len(Trim(meterial)) > 0 Then
    INSDREQUISITN_QR(20) = Null
    INSDREQUISITN_QR(20) = meterial
Else
    INSDREQUISITN_QR(20) = Null
    INSDREQUISITN_QR(20) = ""
End If
INSDREQUISITN_QR(21) = Null
INSDREQUISITN_QR.Execute
'**** INSERTING INTO DREQUISITION

 '*** UPDATING THE PROVISIONAL DIAGNOSIS
 
If opt_patient.Value = True Then
If SELTA1012FLAG = False Then
    selta1012_qr(0) = Msk_caseno.FormattedText
    Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
    SELTA1012FLAG = True
Else
     selta1012_qr(0) = Msk_caseno.FormattedText
     selta1012_rs.MoreResults
     selta1012_rs.Requery
End If
If selta1012_rs.RowCount > 0 Then
    UPDTA1012_QR(0) = Trim(txt_indctinvstg.Text)
    UPDTA1012_QR(1) = Msk_caseno.FormattedText
    UPDTA1012_QR.Execute
Else
    insta1012_qr(0) = Msk_caseno.FormattedText
    insta1012_qr(1) = Trim(txt_indctinvstg.Text)
    insta1012_qr.Execute
End If
End If
'*** UPDATING THE PROVISIONAL DIAGNOSIS
'*** INSERTING INTO DMULPROCE FOR COLLECTION AND DEPT COLLECTION SCREEN
''    INSDMULPROCE_QR(0) = req_chglen(txt_spreqno.Text)
''    INSDMULPROCE_QR(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
''    INSDMULPROCE_QR(2) = Msk_caseno.FormattedText
''    INSDMULPROCE_QR(3) = cmb_procedurecode.Text
''    INSDMULPROCE_QR.Execute

'***INSERTING INTO DMULPROCE FOR COLLECTION AND DEPT COLLECTION SCREEN
'***INSERTING INTO DREQSTAFFD TABLE
If opt_staff.Value = True Then
    INSDREQSTAFFD_QR(0) = req_chglen(txt_bcreqno.Text)
    INSDREQSTAFFD_QR(1) = Msk_caseno.Text
    INSDREQSTAFFD_QR(2) = Trim(txt_name.Text)
    INSDREQSTAFFD_QR(3) = Trim(txt_Age)
    INSDREQSTAFFD_QR(4) = Trim(txt_sex)
    If Len(Trim(cmb_dependents.Text)) = 0 Then
        INSDREQSTAFFD_QR(5) = "SELF"
    Else
        I = InStr(cmb_dependents, "$")
        INSDREQSTAFFD_QR(5) = Mid(cmb_dependents, I + 1, 10)
    End If
    INSDREQSTAFFD_QR.Execute
End If
'***INSERTING INTO DREQSTAFFD TABLE
    If SELCLINHISTFLAG = False Then
        SELCLINHIST_QR(0) = Msk_caseno.FormattedText
        Set SELCLINHIST_RS = SELCLINHIST_QR.OpenResultset(1, 3)
        SELCLINHISTFLAG = True
    Else
        SELCLINHIST_QR(0) = Msk_caseno.FormattedText
        SELCLINHIST_RS.MoreResults
        SELCLINHIST_RS.Requery
    End If
'    If SELCLINHIST_RS.RowCount > 0 Then
'        updclinhist_qr(0) = Trim(txt_clinhist.Text)
'        updclinhist_qr(1) = Trim(Msk_caseno.Text)
'        updclinhist_qr.Execute
'    Else
        insclinhist_qr(0) = Trim(Msk_caseno.FormattedText)
        insclinhist_qr(1) = GET_REVDATE(Trim(DATE11))
        insclinhist_qr(2) = Trim(time1)
        If Trim(txt_clinhist.Text) = "NYD" Then
            insclinhist_qr(3) = "NYD"
        Else
            insclinhist_qr(3) = Trim(txt_clinhist.Text)
        End If
        insclinhist_qr.Execute
    'End If
''UPDATE DREQUISITN TABLE
If SELDISICU_FLAG = False Then
    SELDISICU_QR(0) = Trim(ipadd)
    Set SELDISICU_RS = SELDISICU_QR.OpenResultset(1, 3)
    SELDISICU_FLAG = True
Else
    SELDISICU_QR(0) = Trim(ipadd)
    SELDISICU_RS.MoreResults
    SELDISICU_RS.Requery
End If

If SELDISICU_RS.RowCount > 0 Then
    If Trim(SELDISICU_RS(1)) = "10.3.1.56" Then
        UPDICU_QR(0) = "YR"
        UPDICU_QR(1) = req_chglen(Trim(txt_bcreqno.Text))
        UPDICU_QR.Execute
    Else
        UPDICU_QR(0) = "YES"
        UPDICU_QR(1) = req_chglen(Trim(txt_bcreqno.Text))
        UPDICU_QR.Execute
    End If
End If
'If Not SELDISICU_RS.EOF Then
'    UPDICU_QR(0) = "YES"
'    UPDICU_QR(1) = req_chglen(Trim(txt_bcreqno.Text))
'    UPDICU_QR.Execute
'End If

    
''UPDATE DREQUISITN TABLE
Data1.UpdateRecord
Data1.Recordset.Bookmark = Data1.Recordset.LastModified
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Public Sub MB()
    MBENABL ("ENA")
    surgicalPathology.Visible = False
    immunohistochemistry.Visible = False
    frm_hematopathology.Visible = False
    biochemistry.Visible = False
    bceltm.Visible = False
    List_code.clear
    list_proc.clear
'    cardiology.Visible = False
    Microbiology.Visible = True
    CYTOLOGY.Visible = False
    fnac.Visible = False
    'cpl.Visible = False
    TUMOURMARKER.Visible = False
    hemcbc.Hide
    HEMBF.Hide
    HEMBMA.Hide
   ' HLATYPING.Visible = False
    lbl_req.Left = 3500
    lbl_req.Caption = "REQUISITION - MICROBILOGY"
    pathology.Caption = "MICROBIOLOGY"
    If txt_modcode.Text = "BA" Then
        Microbiology.Caption = "Bacteriology"
        lbl_req.Left = 3500
        lbl_req.Caption = "REQUISITION - BACTERIOLOGY"
        pathology.Caption = "MICROBIOLOGY"
        lbl_mbongant.Visible = True
        txt_ongoinganti.Visible = True
        lbl_mboutslides.Visible = True
        cmb_mboutslide.Visible = True
    End If
    If txt_modcode.Text = "MB" Then
        lbl_req.Left = 3300
        Microbiology.Caption = "MycoBacteriology"
        lbl_req.Caption = "REQUISITION - MYCOBACTERIOLOGY"
        pathology.Caption = "MICROBIOLOGY"
        lbl_mbongant.Visible = True
        txt_ongoinganti.Visible = True
        lbl_mboutslides.Visible = True
        cmb_mboutslide.Visible = True
    End If
    If txt_modcode.Text = "SE" Then
        Microbiology.Caption = "Serology"
        lbl_req.Left = 3600
        lbl_req.Caption = "REQUISITION - SEROLOGY"
        pathology.Caption = "MICROBIOLOGY"
        lbl_mbongant.Visible = False
        txt_ongoinganti.Visible = False
        lbl_mboutslides.Visible = False
        cmb_mboutslide.Visible = False
    End If
    If txt_modcode.Text = "MC" Then
        Microbiology.Caption = "Mycology"
        lbl_req.Left = 3600
        lbl_req.Caption = "REQUISITION - MYCOLOGY"
        pathology.Caption = "MICROBIOLOGY"
        lbl_mbongant.Visible = True
        txt_ongoinganti.Visible = True
        lbl_mboutslides.Visible = True
        cmb_mboutslide.Visible = True
    End If
    If txt_modcode.Text = "CM" Then
        Microbiology.Caption = "Clinical Microbiology"
        lbl_req.Left = 3100
        lbl_req.Caption = "REQUISITION - CLINICAL MICROBIOLOGY"
        pathology.Caption = "MICROBIOLOGY"
        lbl_mbongant.Visible = True
        txt_ongoinganti.Visible = True
        lbl_mboutslides.Visible = False
        cmb_mboutslide.Visible = False
    End If
    If txt_modcode.Text = "MO" Then
        Microbiology.Caption = "Molecular Microbiology"
        lbl_req.Left = 3000
        lbl_req.Caption = "REQUISITION - MOLECULAR MICROBILOGY"
        pathology.Caption = "MICROBIOLOGY"
        lbl_mbongant.Visible = True
        txt_ongoinganti.Visible = True
        lbl_mboutslides.Visible = False
        cmb_mboutslide.Visible = False
    End If
    If txt_modcode.Text = "XX" Then
        Microbiology.Caption = "Others"
        lbl_req.Left = 3700
        lbl_req.Caption = "REQUISITION - OTHERS"
        pathology.Caption = "MICROBIOLOGY"
        lbl_mbongant.Visible = True
        txt_ongoinganti.Visible = True
        lbl_mboutslides.Visible = True
        cmb_mboutslide.Visible = True
    End If
    lbl_mbreqno.Visible = True
    txt_mbreqno.Visible = True
    
    If SELRINVESTFLAG = False Then
        SELRINVEST_QR(0) = Trim(txt_modcode.Text)
        Set SELRINVEST_RS = SELRINVEST_QR.OpenResultset(1, 3)
        SELRINVESTFLAG = True
    Else
        SELRINVEST_QR(0) = Trim(txt_modcode.Text)
        SELRINVEST_RS.MoreResults
        SELRINVEST_RS.Requery
    End If
    If SELRINVEST_RS.RowCount > 0 Then
        list_proc.clear
        List_code.clear
        I1 = 0
        While Not SELRINVEST_RS.EOF
            List_code.AddItem Trim(SELRINVEST_RS("ser_code"))
            If SELMA101FLAG = False Then
                SELMA101_QR(0) = Trim(SELRINVEST_RS("SER_CODE"))
                Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                SELMA101FLAG = True
             Else
                SELMA101_QR(0) = SELRINVEST_RS("SER_CODE")
                SELMA101_RS.MoreResults
                SELMA101_RS.Requery
            End If
            
            If SELMA101_RS.RowCount > 0 Then
                While Not SELMA101_RS.EOF
                                               
                list_proc.AddItem SELMA101_RS("SER_DET")
                             
                    If Trim(SELMA101_RS("SER_DET")) = "ADA LEVEL" Then
                        g253cd = I1
                    End If
                    If Trim(SELMA101_RS("SER_DET")) = "CULTURE & SENSITIVITY (AFB)" Then
                        g106cd = I1
                    End If
                    If Trim(SELMA101_RS("SER_DET")) = "AFB CULTURE ONLY" Then
                        g119cd = I1
                    End If
                    
                SELMA101_RS.MoveNext
                I1 = I1 + 1
                Wend
            End If
            SELRINVEST_RS.MoveNext
        Wend
    End If
    
    ' PROCEDURE CODE FILLING IN THE COMBOBOX
              If seLdreqrsflag = False Then
                    SELDREQQR(0) = Msk_caseno.FormattedText
                    Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                    seLdreqrsflag = True
                Else
                    SELDREQQR(0) = Msk_caseno.FormattedText
                    SELDREQRS.MoreResults
                    SELDREQRS.Requery
                End If
                If SELDREQRS.RowCount > 0 Then
                    cmb_mbprevunv.clear
                    While Not SELDREQRS.EOF
                       cmb_mbprevunv.AddItem SELDREQRS("PATHNO") & "$" & SELDREQRS("REQNO")
                        SELDREQRS.MoveNext
                    Wend
                    cmb_mbprevunv.ListIndex = 0
                End If
    
    For I = 0 To 10
        cmb_mboutslide.AddItem I
    Next I
    
End Sub

Public Sub MBASSIGN()
Dim durill As String
Dim I As Integer

On Error GoTo MSG:
'*** INSERTING INT0 RREQUISITION
INSRREQUISITN_QR(0) = req_chglen(Trim(txt_mbreqno.Text))
INSRREQUISITN_QR(1) = GET_REVDATE(CStr(DATE11))
INSRREQUISITN_QR(2) = Trim(txt_time.Text)
INSRREQUISITN_QR(3) = Trim(Msk_caseno.FormattedText)
INSRREQUISITN_QR(4) = Trim(txt_dept.Text)
INSRREQUISITN_QR(5) = Trim(txt_modcode.Text)
INSRREQUISITN_QR(6) = Trim(txt_status.Text)
INSRREQUISITN_QR(7) = Trim(txt_catgory.Text)
INSRREQUISITN_QR(8) = "8888"
INSRREQUISITN_QR(9) = UCase(Trim(cmb_refdoc.Text))
'INSRREQUISITN_QR(9) = Trim(txt_pridocser.Text)

'Added by Milind on 27-03-2006 for inserting date for the samples collected at outside
    If Check1.Value = 1 Then
        Dim Dt As String
        Dt = "18/02/1980"
        INSRREQUISITN_QR(15) = Format(CDate(Dt), "dd/mm/yyyy")
    Else
        INSRREQUISITN_QR(15) = Null
    End If
'upto here

    If SELRREQUISITNFLAG = False Then
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
        SELRREQUISITNFLAG = True
    Else
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        SELRREQUISITN_RS.MoreResults
        SELRREQUISITN_RS.Requery
    End If
    If SELRREQUISITN_RS.RowCount > 0 Then
        INSRREQUISITN_QR(10) = "YES"
    Else
        INSRREQUISITN_QR(10) = "NO"
    End If
If opt_patient.Value = True Then
    NO_CHARGE
    ''' Added by Sandhya dated 26.6.2009
    ''' for NC category & charged patient
    If txt_catgory.Text = "NC" And No_Chargeflag = True Then
        INSRREQUISITN_QR(11) = "NOCHARGE"
        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
        INSRREQUISITN_QR(13) = "P"
        'No_Chargeflag = False
        
    ''ICU
    ElseIf lbl_icureq.Visible = True Then
        'If Trim(CMB_ICU1.Text) = "YES" Then
            INSRREQUISITN_QR(11) = Null
            INSRREQUISITN_QR(12) = Null
           ' INSRREQUISITN_QR(13) = "P"
            'INSRREQUISITN_QR(13) = "E"
           ' INSRREQUISITN_QR(13) = " "
             INSRREQUISITN_QR(13) = Null
        'End If
    ''ICU
    Else
        
            INSRREQUISITN_QR(11) = Null
            INSRREQUISITN_QR(12) = Null
            INSRREQUISITN_QR(13) = Null
       
    End If
End If
If opt_staff.Value = True Then
        INSRREQUISITN_QR(11) = "NOCHARGE"
        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
        INSRREQUISITN_QR(13) = "P"
End If
INSRREQUISITN_QR(14) = 0

INSRREQUISITN_QR.Execute
    '*** INSERTING INT0 RREQUISITION
    '*** INSERTING INTO DMULPROCE TABLE
    If List_code.ListCount <> 0 Then
        For I = 0 To List_code.ListCount - 1
            If List_code.Selected(I) = True Then
                INSDMULPROCE_QR(0) = req_chglen(txt_mbreqno.Text)
                INSDMULPROCE_QR(1) = List_code.List(I)
                INSDMULPROCE_QR(2) = GET_REVDATE(DATE11)
                INSDMULPROCE_QR(3) = Msk_caseno.FormattedText
                INSDMULPROCE_QR.Execute
            End If
        Next I
    End If
    '*** INSERTING INTO DMULPROCE TABLE
    '*** INSERTING INTO DMSPONANTIB
        INSSPONANTIBQR(0) = req_chglen(txt_mbreqno.Text)
        'If Len(Trim(txt_mbspecimen.Text)) > 0 Then
        If Len(Trim(cmb_specimen.Text)) > 0 Then
            INSSPONANTIBQR(1) = Null
            'INSSPONANTIBQR(1) = txt_mbspecimen
            INSSPONANTIBQR(1) = cmb_specimen.Text
        Else
            INSSPONANTIBQR(1) = Null
            INSSPONANTIBQR(1) = " "
        End If
        If Len(Trim(txt_ongoinganti.Text)) > 0 Then
            INSSPONANTIBQR(2) = Null
            INSSPONANTIBQR(2) = txt_ongoinganti.Text
        Else
            INSSPONANTIBQR(2) = Null
            INSSPONANTIBQR(2) = " "
        End If
        INSSPONANTIBQR.Execute
    
    '*** INSERTING INTO DMSPONANTIB
    '**** INSERTING INTO DREQUISITION

INSDREQUISITN_QR(0) = req_chglen(Trim(txt_mbreqno.Text))
INSDREQUISITN_QR(1) = TXT_MBPATHNO.Text
INSDREQUISITN_QR(2) = Trim(Msk_caseno.FormattedText)
                            
' i = InStr(cmb_sites, "$")
' If i = 0 Then
'
'               If sitedupflag = False Then
'                    sitedupqr(0) = UCase(Trim(cmb_sites.Text))
'                    Set siteduprs = sitedupqr.OpenResultset(1, 3)
'                    sitedupflag = True
'               Else
'                    sitedupqr(0) = UCase(Trim(cmb_sites.Text))
'                    siteduprs.MoreResults
'                    siteduprs.Requery
'
'               End If
'
'        If siteduprs.RowCount > 0 Then
'                INSDREQUISITN_QR(3) = siteduprs("SITECD")
'        Else
'                If sitedup1flag = False Then
'                    'sitedupqr(0) = Trim(cmb_sites.Text)
'                    Set siteduprs1 = sitedupqr1.OpenResultset(1, 3)
'                    sitedup1flag = True
'                Else
'                    siteduprs1.MoreResults
'                    siteduprs1.Requery
'                End If
'                If (siteduprs1.RowCount > 0) Then
'                    Dim A As Integer
'                    If Len(Trim(cmb_sites.Text)) > 0 Then
'
'                    A = siteduprs1(0)
'                    A = A + 1
'                    siteinsertqr(0) = A
'                    siteinsertqr(1) = UCase(Trim(cmb_sites.Text))
'                    siteinsertqr.Execute
'                    INSDREQUISITN_QR(3) = A
'                    Else
'                        INSDREQUISITN_QR(3) = "9999"
'                    End If
'                 Else
'
'                    A = 5000
'                    siteinsertqr(0) = A
'                    siteinsertqr(1) = UCase(Trim(cmb_sites.Text))
'                    siteinsertqr.Execute
'                    INSDREQUISITN_QR(3) = A
'                 End If
'         End If
'  Else
        INSDREQUISITN_QR(3) = Null
'  End If
INSDREQUISITN_QR(4) = Null
dur_ill = txt_year & "/" & Trim(CMB_DURILLMONTH.Text)
INSDREQUISITN_QR(5) = dur_ill
INSDREQUISITN_QR(6) = Null
INSDREQUISITN_QR(7) = Null
INSDREQUISITN_QR(8) = Null
INSDREQUISITN_QR(9) = Null
INSDREQUISITN_QR(10) = cmb_mboutslide.Text
INSDREQUISITN_QR(11) = Null
INSDREQUISITN_QR(12) = Null
INSDREQUISITN_QR(13) = Null
INSDREQUISITN_QR(14) = Trim(cmb_refdoc.Text)
INSDREQUISITN_QR(15) = Trim(cmb_reftodoc.Text)
INSDREQUISITN_QR(16) = Null
If Trim(txt_status.Text) = "IP" Then
    INSDREQUISITN_QR(17) = Trim(txt_ward.Text)
    INSDREQUISITN_QR(18) = Trim(txt_bedno.Text)
Else
    INSDREQUISITN_QR(17) = Null
    INSDREQUISITN_QR(18) = Null
End If
INSDREQUISITN_QR(19) = " "
INSDREQUISITN_QR(20) = Null
INSDREQUISITN_QR(21) = Null 'CMB_ICU.Text
INSDREQUISITN_QR.Execute
'**** INSERTING INTO DREQUISITION

 '*** UPDATING THE PROVISIONAL DIAGNOSIS
 
If opt_patient.Value = True Then
If SELTA1012FLAG = False Then
    selta1012_qr(0) = Msk_caseno.FormattedText
    Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
    SELTA1012FLAG = True
Else
     selta1012_qr(0) = Msk_caseno.FormattedText
     selta1012_rs.MoreResults
     selta1012_rs.Requery
End If
If selta1012_rs.RowCount > 0 Then
    UPDTA1012_QR(0) = Trim(txt_indctinvstg.Text)
    UPDTA1012_QR(1) = Msk_caseno.FormattedText
    UPDTA1012_QR.Execute
Else
    insta1012_qr(0) = Msk_caseno.FormattedText
    insta1012_qr(1) = Trim(txt_indctinvstg.Text)
    insta1012_qr.Execute
End If
End If
'*** UPDATING THE PROVISIONAL DIAGNOSIS
'*** INSERTING INTO DMULPROCE FOR COLLECTION AND DEPT COLLECTION SCREEN
''    INSDMULPROCE_QR(0) = req_chglen(txt_spreqno.Text)
''    INSDMULPROCE_QR(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
''    INSDMULPROCE_QR(2) = Msk_caseno.FormattedText
''    INSDMULPROCE_QR(3) = cmb_procedurecode.Text
''    INSDMULPROCE_QR.Execute

'***INSERTING INTO DMULPROCE FOR COLLECTION AND DEPT COLLECTION SCREEN
'***INSERTING INTO DREQSTAFFD TABLE
If opt_staff.Value = True Then
    INSDREQSTAFFD_QR(0) = req_chglen(txt_mbreqno.Text)
    INSDREQSTAFFD_QR(1) = Msk_caseno.Text
    INSDREQSTAFFD_QR(2) = Trim(txt_name.Text)
    INSDREQSTAFFD_QR(3) = Trim(txt_Age)
    INSDREQSTAFFD_QR(4) = Trim(txt_sex)
    If Len(Trim(cmb_dependents.Text)) = 0 Then
        INSDREQSTAFFD_QR(5) = "SELF"
    Else
        I = InStr(cmb_dependents, "$")
        INSDREQSTAFFD_QR(5) = Mid(cmb_dependents, I + 1, 10)
    End If
    INSDREQSTAFFD_QR.Execute
End If
'***INSERTING INTO DREQSTAFFD TABLE
    If SELCLINHISTFLAG = False Then
        SELCLINHIST_QR(0) = Msk_caseno.FormattedText
        Set SELCLINHIST_RS = SELCLINHIST_QR.OpenResultset(1, 3)
        SELCLINHISTFLAG = True
    Else
        SELCLINHIST_QR(0) = Msk_caseno.FormattedText
        SELCLINHIST_RS.MoreResults
        SELCLINHIST_RS.Requery
    End If
'    If SELCLINHIST_RS.RowCount > 0 Then
'        updclinhist_qr(0) = Trim(txt_clinhist.Text)
'        updclinhist_qr(1) = Trim(Msk_caseno.Text)
'        updclinhist_qr.Execute
'    Else
        insclinhist_qr(0) = Trim(Msk_caseno.FormattedText)
        insclinhist_qr(1) = GET_REVDATE(Trim(DATE11))
        insclinhist_qr(2) = Trim(time1)
        If Trim(txt_clinhist.Text) = "NYD" Then
            insclinhist_qr(3) = "NYD"
        Else
            insclinhist_qr(3) = Trim(txt_clinhist.Text)
        End If
        insclinhist_qr.Execute
    'End If
''UPDATE DREQUISITN TABLE
If SELDISICU_FLAG = False Then
    SELDISICU_QR(0) = Trim(ipadd)
    Set SELDISICU_RS = SELDISICU_QR.OpenResultset(1, 3)
    SELDISICU_FLAG = True
Else
    SELDISICU_QR(0) = Trim(ipadd)
    SELDISICU_RS.MoreResults
    SELDISICU_RS.Requery
End If
If Not SELDISICU_RS.EOF Then
    UPDICU_QR(0) = "YES"
    UPDICU_QR(1) = req_chglen(Trim(txt_mbreqno.Text))
    UPDICU_QR.Execute
End If
    
''UPDATE DREQUISITN TABLE
Data1.UpdateRecord
Data1.Recordset.Bookmark = Data1.Recordset.LastModified
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Public Sub MBCLEAR()
    List_code.clear
    list_proc.clear
    cmb_mboutslide.Text = ""
    txt_ongoinganti.Text = ""
    'txt_mbspecimen.Text = ""
    cmb_specimen.Text = ""
    'CMB_ICU.Text = ""
End Sub

Public Sub MBENABL(EN As String)
If EN = "DIS" Then
        List_code.Enabled = False
        list_proc.Enabled = False
        cmb_mboutslide.Enabled = False
        'txt_mbspecimen.Enabled = False
        cmb_specimen.Enabled = False
        txt_ongoinganti.Enabled = False
        cmb_mbprevunv.Enabled = False
        CMB_ICU.Enabled = False
        
End If
If EN = "ENA" Then
        List_code.Enabled = True
        list_proc.Enabled = True
        cmb_mboutslide.Enabled = True
        'txt_mbspecimen.Enabled = True
        cmb_specimen.Enabled = True
        txt_ongoinganti.Enabled = True
        cmb_mbprevunv.Enabled = True
        CMB_ICU.Enabled = True
End If
If EN = "MOD" Then
        List_code.Enabled = False
        list_proc.Enabled = False
        cmb_mboutslide.Enabled = True
        'txt_mbspecimen.Enabled = True
        cmb_specimen.Enabled = True
        txt_ongoinganti.Enabled = True
        cmb_mbprevunv.Enabled = True
        CMB_ICU.Enabled = True
End If
End Sub

Public Sub MBUPD()
'**** Updating the drequisition table
        upddrequisitn_qr(0) = (cmb_refdoc.Text)
        upddrequisitn_qr(1) = cmb_reftodoc.Text
        upddrequisitn_qr(2) = txt_year & "/" & CMB_DURILLMONTH
        upddrequisitn_qr(3) = Null
        upddrequisitn_qr(4) = Null
        upddrequisitn_qr(5) = Null
        upddrequisitn_qr(6) = Null
        upddrequisitn_qr(7) = Null
        upddrequisitn_qr(8) = Null
        upddrequisitn_qr(9) = cmb_mboutslide.Text
        upddrequisitn_qr(10) = Null
        upddrequisitn_qr(11) = Null
        upddrequisitn_qr(12) = "NO"
        upddrequisitn_qr(13) = Null
        upddrequisitn_qr(14) = Null
        upddrequisitn_qr(15) = req_chglen(cmb_reqno.Text)
        upddrequisitn_qr.Execute
        '**** Updating the drequisition table
        'If Len(Trim(txt_mbspecimen.Text)) > 0 Then
        If Len(Trim(cmb_specimen.Text)) > 0 Then
            UPDSPONANTIBQR(0) = Null
            'UPDSPONANTIBQR(0) = txt_mbspecimen.Text
            UPDSPONANTIBQR(0) = cmb_specimen.Text
        Else
            UPDSPONANTIBQR(0) = Null
            UPDSPONANTIBQR(0) = " "
        End If
        If Len(Trim(txt_ongoinganti.Text)) > 0 Then
            UPDSPONANTIBQR(1) = Null
            UPDSPONANTIBQR(1) = txt_ongoinganti.Text
        Else
            UPDSPONANTIBQR(1) = Null
            UPDSPONANTIBQR(1) = " "
        End If
        UPDSPONANTIBQR(2) = req_chglen(cmb_reqno.Text)
        UPDSPONANTIBQR.Execute
        
        
         '*** UPDATING THE PROVISIONAL DIAGNOSIS
        If opt_patient.Value = True Then
        If SELTA1012FLAG = False Then
            selta1012_qr(0) = Msk_caseno.FormattedText
            Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
            SELTA1012FLAG = True
        Else
             selta1012_qr(0) = Msk_caseno.FormattedText
             selta1012_rs.MoreResults
             selta1012_rs.Requery
        End If
        If selta1012_rs.RowCount > 0 Then
            UPDTA1012_QR(0) = Trim(txt_indctinvstg.Text)
            UPDTA1012_QR(1) = Msk_caseno.FormattedText
            UPDTA1012_QR.Execute
        Else
            insta1012_qr(0) = Msk_caseno.FormattedText
            insta1012_qr(1) = Trim(txt_indctinvstg.Text)
            insta1012_qr.Execute
        End If
        End If
        '*** UPDATING THE PROVISIONAL DIAGNOSIS
        '*** updating the clinical history
        If SELCLINHISTFLAG = False Then
            SELCLINHIST_QR(0) = Msk_caseno.FormattedText
            Set SELCLINHIST_RS = SELCLINHIST_QR.OpenResultset(1, 3)
            SELCLINHISTFLAG = True
        Else
            SELCLINHIST_QR(0) = Msk_caseno.FormattedText
            SELCLINHIST_RS.MoreResults
            SELCLINHIST_RS.Requery
        End If
        'If SELCLINHIST_RS.RowCount > 0 Then
'            updclinhist_qr(0) = Trim(txt_clinhist.Text)
'            updclinhist_qr(1) = Trim(Msk_caseno.FormattedText)
'            updclinhist_qr.Execute
'        Else
            insclinhist_qr(0) = Trim(Msk_caseno.FormattedText)
            insclinhist_qr(1) = GET_REVDATE(Trim(DATE11))
            insclinhist_qr(2) = Trim(time1)
            If Trim(txt_clinhist.Text) = "NYD" Then
                insclinhist_qr(3) = "NYD"
            Else
                insclinhist_qr(3) = Trim(txt_clinhist.Text)
            End If
            insclinhist_qr.Execute
        'End If
        '*** UPDATING THE RREQUISITN TABLE
        updrrequisitn_qr(0) = UCase(Trim(cmb_refdoc.Text))
        updrrequisitn_qr(1) = req_chglen(cmb_reqno.Text)
        updrrequisitn_qr.Execute
     ' End If

       
        '*** UPDATING THE RREQUISITN TABLE
End Sub

Public Sub TMVIEW()
    TUMOURMARKER.Msk_caseno.Text = Msk_caseno.FormattedText
    TUMOURMARKER.txt_name = txt_name.Text
   
        If SELDMULPROCERS("PRO_CODE") = "F818" Then
            TUMOURMARKER.serum_CEA.Value = 1
         End If
        
        'ADDED BY MILIND
        If SELDMULPROCERS("PRO_CODE") = "F824" Then
            TUMOURMARKER.serumCA19_9.Value = 1
         End If
        'DATE 09/01/06
         
        If SELDMULPROCERS("PRO_CODE") = "F817" Then
            TUMOURMARKER.serum_AFP.Value = 1
         End If
    
    
        If SELDMULPROCERS("PRO_CODE") = "F819" Then
            TUMOURMARKER.serumB_HCG.Value = 1
         End If
         
        
        If SELDMULPROCERS("PRO_CODE") = "F823" Then
            TUMOURMARKER.serum_CA_125.Value = 1
        End If
        
   
        If SELDMULPROCERS("PRO_CODE") = "F820" Then
            TUMOURMARKER.serum_PSA.Value = 1
         End If
      
   
        If SELDMULPROCERS("PRO_CODE") = "F821" Then
           ' TUMOURMARKER.serumB2_Microglobin.Value = 1
            biochemistry.serumB2_Microglobin.Value = 1
                       'shweta
         End If
        
   
        If SELDMULPROCERS("PRO_CODE") = "F826" Then
            'TUMOURMARKER.Serum_PLAP.Value = 1
         End If
       
    
        If SELDMULPROCERS("PRO_CODE") = "F822" Then
            TUMOURMARKER.serumCA_153.Value = 1
         End If
      
   
        If SELDMULPROCERS("PRO_CODE") = "F816" Then
            TUMOURMARKER.ImmunoelectrophoresisIEP.Value = 1
         End If
        
        If SELDMULPROCERS("PRO_CODE") = "F832" Then
            TUMOURMARKER.CHK_NSE.Value = 1
         End If
        
        If SELDMULPROCERS("PRO_CODE") = "F830" Then
            TUMOURMARKER.serumFerritin.Value = 1
       
        End If
        
        If SELDMULPROCERS("PRO_CODE") = "F831" Then
            TUMOURMARKER.CHK_CYFRA.Value = 1
       
        End If
    
        If SELDMULPROCERS("PRO_CODE") = "F833" Then
            TUMOURMARKER.chk_cyclos.Value = 1
       
        End If
        
    
    
        If SELDMULPROCERS("PRO_CODE") = "F834" Then
            TUMOURMARKER.chk_amik.Value = 1
        
        End If
        
    
        If SELDMULPROCERS("PRO_CODE") = "F835" Then
            TUMOURMARKER.chk_vanc.Value = 1
        
        End If
    
        If SELDMULPROCERS("PRO_CODE") = "F836" Then
            TUMOURMARKER.chk_metho.Value = 1
        
        End If
      
        If SELDMULPROCERS("PRO_CODE") = "F828" Then
            TUMOURMARKER.chk_bjp.Value = 1
        
        End If
     
    
        If SELDMULPROCERS("PRO_CODE") = "F827" Then
            TUMOURMARKER.chk_addmark.Value = 1
        End If
       
End Sub

Public Sub ELVIEW()
'If SELDMULPROCERS("pro_coDe") = "F860" Then
If SELDMULPROCERS("pro_coDe") = "F981" Then
    Em_lab.chk_elect.Value = 1
End If
If SELDMULPROCERS("pro_coDe") = "F995" Then
    Em_lab.chk_venous.Value = 1
End If
If SELDMULPROCERS("pro_coDe") = "F998" Then
    Em_lab.chk_osmolality.Value = 1
End If
If SELDMULPROCERS("pro_coDe") = "F022" Then
    Em_lab.chk_bgel.Value = 1
End If

If SELDMULPROCERS("pro_coDe") = "F986" Then
    Em_lab.chk_uricacid.Value = 1
End If
If SELDMULPROCERS("pro_coDe") = "F875" Then
    Em_lab.chk_cardenz.Value = 1
End If
'If SELDMULPROCERS("PRO_CODE") = "F842" Then
If SELDMULPROCERS("PRO_CODE") = "F979" Then
    Em_lab.chk_FastSug.Value = 1
End If
'If SELDMULPROCERS("PRO_CODE") = "F843" Then
If SELDMULPROCERS("PRO_CODE") = "F980" Then
     Em_lab.chk_ppsug.Value = 1
End If
'If SELDMULPROCERS("PRO_CODE") = "F841" Then
If SELDMULPROCERS("PRO_CODE") = "F978" Then
    Em_lab.chk_ransug.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "F848" Then
     Em_lab.chk_gluco.Value = 1
End If
'If SELDMULPROCERS("PRO_CODE") = "F855" Then
If SELDMULPROCERS("PRO_CODE") = "F985" Then
    Em_lab.chk_urea.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "F856" Then
    Em_lab.chk_creat.Value = 1
End If
'If SELDMULPROCERS("PRO_CODE") = "F857" Then
If SELDMULPROCERS("PRO_CODE") = "F987" Then
    Em_lab.chk_creat.Value = 1
End If
'If SELDMULPROCERS("PRO_CODE") = "F863" Then
If SELDMULPROCERS("PRO_CODE") = "F984" Then
     Em_lab.chk_chlor.Value = 1
End If
'If SELDMULPROCERS("PRO_CODE") = "F861" Then
If SELDMULPROCERS("PRO_CODE") = "F982" Then
    Em_lab.chk_sodium.Value = 1
End If
'If SELDMULPROCERS("PRO_CODE") = "F862" Then
If SELDMULPROCERS("PRO_CODE") = "F983" Then
     Em_lab.chk_pota.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "F877" Then
    Em_lab.chk_sgot.Value = 1
 End If
'If SELDMULPROCERS("PRO_CODE") = "F872" Then
'    Em_lab.chk_sgpt.Value = 1
' End If
If SELDMULPROCERS("PRO_CODE") = "F879" Then
    Em_lab.chk_cpkmb.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "F876" Then
    Em_lab.chk_ldh.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "F878" Then
    Em_lab.CHK_cpk.Value = 1
End If
'If SELDMULPROCERS("PRO_CODE") = "F891" Then
 If SELDMULPROCERS("PRO_CODE") = "F993" Then
    Em_lab.chk_magn.Value = 1
End If
'If SELDMULPROCERS("PRO_CODE") = "F892" Then
If SELDMULPROCERS("PRO_CODE") = "F990" Then
    Em_lab.chk_lactate.Value = 1
 End If
'If SELDMULPROCERS("PRO_CODE") = "F881" Then
If SELDMULPROCERS("PRO_CODE") = "F988" Then
    Em_lab.chk_amylase.Value = 1
End If
'If SELDMULPROCERS("PRO_CODE") = "F888" Then
 If SELDMULPROCERS("PRO_CODE") = "F991" Then
    Em_lab.chk_ca1.Value = 1
End If
'If SELDMULPROCERS("PRO_CODE") = "F889" Then
If SELDMULPROCERS("PRO_CODE") = "F992" Then
    Em_lab.chk_ca.Value = 1
End If
'''*** ADD BY MONO ON 09/04/2003

If SELDMULPROCERS("pro_code") = "F901" Then
    Em_lab.chk_am.Value = 1
End If
If SELDMULPROCERS("pro_code") = "F902" Then
    Em_lab.chk_cap.Value = 1
End If

If SELDMULPROCERS("pro_code") = "F903" Then
    Em_lab.CHK_hbpcv.Value = 1
End If

If SELDMULPROCERS("pro_code") = "F904" Then
    Em_lab.CHK_HB.Value = 1
End If
If SELDMULPROCERS("pro_code") = "F906" Then
    Em_lab.CHK_PC.Value = 1
End If
If SELDMULPROCERS("pro_code") = "F907" Then
    Em_lab.chk_pcv.Value = 1
End If


If SELDMULPROCERS("pro_code") = "F908" Then
    Em_lab.chk_btct.Value = 1
End If

If SELDMULPROCERS("pro_code") = "F909" Then
    Em_lab.CHK_BT.Value = 1
End If
If SELDMULPROCERS("pro_code") = "F910" Then
    Em_lab.CHK_CT.Value = 1
End If


If SELDMULPROCERS("pro_code") = "F911" Then
    Em_lab.chk_pt.Value = 1
End If
If SELDMULPROCERS("pro_code") = "F912" Then
    Em_lab.CHK_PTT.Value = 1
End If
If SELDMULPROCERS("pro_code") = "F913" Then
    Em_lab.CHK_TT.Value = 1
End If
If SELDMULPROCERS("pro_code") = "F905" Then
    Em_lab.Chk_wbc.Value = 1
End If
'LFT
If SELDMULPROCERS("PRO_CODE") = "F949" Then
    Em_lab.chk_livfuntest.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "F950" Then
    Em_lab.chk_prot.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "F955" Then
    Em_lab.chk_albumin.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "F956" Then
    Em_lab.chk_glob.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "F957" Then
    Em_lab.chk_alkphos.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "F958" Then
    Em_lab.chk_totbil.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "F959" Then
    Em_lab.chk_dirbil.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "F960" Then
    Em_lab.chk_indirbil.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "F877" Then
    Em_lab.Chk_SGOT1.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "F961" Then
    Em_lab.chk_sgpt.Value = 1
End If
End Sub

Public Sub BCVIEW()
'' CCT CODE ADDED
'HEREE
If SELDMULPROCERS("PRO_CODE") = "F924" Then
      BIOCHEMISTRY1.CHK_CCC.Value = 1
      BIOCHEMISTRY1.CHK_SC.Value = 1
      BIOCHEMISTRY1.CHK_UC.Value = 1
      BIOCHEMISTRY1.GROUPCCT.Value = 1
      BIOCHEMISTRY1.chk_corrcreatclr24.Value = 1
      'BIOCHEMISTRY1.GROUPCCT.Value = 1
      GoTo quit
 End If
If SELDMULPROCERS("PRO_CODE") = "F920" Then
      BIOCHEMISTRY1.chk_creat24.Value = 1
      GoTo quit
 End If
 
If SELDMULPROCERS("PRO_CODE") = "F999" Then
      BIOCHEMISTRY1.chk_ggt.Value = 1
      GoTo quit
End If

If SELDMULPROCERS("PRO_CODE") = "F974" Then
      BIOCHEMISTRY1.CHK_SC.Value = 1
      GoTo quit
 End If
'''URINE RANDOM         1
If SELDMULPROCERS("PRO_CODE") = "F925" Then
      BIOCHEMISTRY1.chk_ureaur.Value = 1
      GoTo quit
 End If
If SELDMULPROCERS("PRO_CODE") = "F926" Then
     BIOCHEMISTRY1.chk_uricacidur.Value = 1
     GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F927" Then
      BIOCHEMISTRY1.chk_creatur.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F845" Then
     BIOCHEMISTRY1.chk_glyhb.Value = 1
     GoTo quit
End If
''Na, k,Cl
If SELDMULPROCERS("PRO_CODE") = "F928" Then
     BIOCHEMISTRY1.chk_sodiumur.Value = 1
     GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F930" Then
      BIOCHEMISTRY1.chk_chlorur.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F929" Then
      BIOCHEMISTRY1.chk_potaur.Value = 1
      GoTo quit
End If
''Ca, P
If SELDMULPROCERS("PRO_CODE") = "F931" Then
      BIOCHEMISTRY1.chk_calcur.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F932" Then
     BIOCHEMISTRY1.chk_phospur.Value = 1
     GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F971" Then
     BIOCHEMISTRY1.CHK_AMYBF.Value = 1
     GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F972" Then
     BIOCHEMISTRY1.CHK_LIPBF.Value = 1
     GoTo quit
End If

''Protein
If SELDMULPROCERS("PRO_CODE") = "F933" Then
     BIOCHEMISTRY1.chk_protur.Value = 1
     GoTo quit
End If
''***urinary investigation
If SELDMULPROCERS("PRO_CODE") = "F899" Then
     BIOCHEMISTRY1.chk_urobur.Value = 1
     GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F900" Then
     BIOCHEMISTRY1.chk_Porphyrinur.Value = 1
     GoTo quit
End If
'''URINE RANDOM             1
'****************************************************

'****************************************************
'''24 HOURS URINE            2
If SELDMULPROCERS("PRO_CODE") = "F918" Then
      BIOCHEMISTRY1.chk_urea24.Value = 1
      GoTo quit
End If

If SELDMULPROCERS("PRO_CODE") = "F919" Then
     BIOCHEMISTRY1.chk_uricacid24.Value = 1
     GoTo quit
End If

If SELDMULPROCERS("PRO_CODE") = "F920" Then
      BIOCHEMISTRY1.chk_creat24.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F924" Then
      BIOCHEMISTRY1.chk_corrcreatclr24.Value = 1
      GoTo quit
End If

''Na, k,Cl
If SELDMULPROCERS("PRO_CODE") = "F915" Then
     BIOCHEMISTRY1.chk_sodium24.Value = 1
     GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F917" Then
      BIOCHEMISTRY1.chk_chlor24.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F916" Then
      BIOCHEMISTRY1.chk_pota24.Value = 1
      GoTo quit
End If
''Ca, P
If SELDMULPROCERS("PRO_CODE") = "F921" Then
      BIOCHEMISTRY1.chk_calc24.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F922" Then
     BIOCHEMISTRY1.chk_phosp24.Value = 1
     GoTo quit
End If
''Protein
If SELDMULPROCERS("PRO_CODE") = "F923" Then
     BIOCHEMISTRY1.chk_prot24.Value = 1
     GoTo quit
End If

'''URINE 24 HOURS         2
'***************************************************
'3
'*************************************************
If SELDMULPROCERS("PRO_CODE") = "F802" Then
    BIOCHEMISTRY1.chk_routinebio.Value = 1
    GoTo quit
End If

If SELDMULPROCERS("PRO_CODE") = "F842" Then
'If SELDMULPROCERS("PRO_CODE") = "F979" Then
      BIOCHEMISTRY1.chk_trusugfas.Value = 1
      GoTo quit
End If

If SELDMULPROCERS("PRO_CODE") = "F843" Then
'If SELDMULPROCERS("PRO_CODE") = "F980" Then
      BIOCHEMISTRY1.chk_ppsug.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F841" Then
'If SELDMULPROCERS("PRO_CODE") = "F978" Then
      BIOCHEMISTRY1.chk_ransug.Value = 1
      GoTo quit
End If
'If SELDMULPROCERS("PRO_CODE") = "F841" Then
'      biochemistry.chk_glucosug.Value = 1
'End If
If SELDMULPROCERS("PRO_CODE") = "F846" Then
      BIOCHEMISTRY1.chk_furinsug.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F847" Then
      BIOCHEMISTRY1.chk_ppurinsug.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "845" Then
      BIOCHEMISTRY1.chk_glyhb.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F849" Then
      BIOCHEMISTRY1.chk_lippro.Value = 1
      BIOCHEMISTRY1.chk_cholest.Value = 1
      BIOCHEMISTRY1.chk_hdlcholest.Value = 1
      BIOCHEMISTRY1.chk_trigly.Value = 1
      BIOCHEMISTRY1.chk_ldlcholest.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F850" Then
      BIOCHEMISTRY1.chk_cholest.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F851" Then
      BIOCHEMISTRY1.chk_hdlcholest.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F853" Then
      BIOCHEMISTRY1.chk_trigly.Value = 1
      GoTo quit
 End If
If SELDMULPROCERS("PRO_CODE") = "F852" Then
      BIOCHEMISTRY1.chk_ldlcholest.Value = 1
      GoTo quit
 End If
If SELDMULPROCERS("PRO_CODE") = "F854" Then
    BIOCHEMISTRY1.chk_renfuntest.Value = 1
    GoTo quit
End If

If SELDMULPROCERS("PRO_CODE") = "F855" Then
'If SELDMULPROCERS("PRO_CODE") = "F985" Then
      BIOCHEMISTRY1.chk_urea.Value = 1
      GoTo quit
 End If
If SELDMULPROCERS("PRO_CODE") = "F856" Then
     BIOCHEMISTRY1.chk_uricacid.Value = 1
     GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F857" Then
'If SELDMULPROCERS("PRO_CODE") = "F987" Then
      BIOCHEMISTRY1.chk_creat.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F859" Then
    ' BIOCHEMISTRY1.chk_corrcreatclr.Value = 1
    ' GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F860" Then
'If SELDMULPROCERS("PRO_CODE") = "F981" Then
    BIOCHEMISTRY1.chk_elect.Value = 1
    GoTo quit
End If

If SELDMULPROCERS("PRO_CODE") = "F861" Then
'If SELDMULPROCERS("PRO_CODE") = "F982" Then
     BIOCHEMISTRY1.chk_sodium.Value = 1
     GoTo quit
End If

If SELDMULPROCERS("PRO_CODE") = "F863" Then
'If SELDMULPROCERS("PRO_CODE") = "F984" Then
      BIOCHEMISTRY1.chk_chlor.Value = 1
      GoTo quit
End If

If SELDMULPROCERS("PRO_CODE") = "F862" Then
'If SELDMULPROCERS("PRO_CODE") = "F983" Then
      BIOCHEMISTRY1.chk_pota.Value = 1
      GoTo quit
End If
 
If SELDMULPROCERS("PRO_CODE") = "F864" Then
     BIOCHEMISTRY1.CHK_CO2.Value = 1
     GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F888" Then
'If SELDMULPROCERS("PRO_CODE") = "F991" Then
      BIOCHEMISTRY1.chk_calc.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F890" Then
     BIOCHEMISTRY1.chk_phosp.Value = 1
     GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F889" Then
'If SELDMULPROCERS("PRO_CODE") = "F992" Then
      BIOCHEMISTRY1.chk_ioncal.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F891" Then
'If SELDMULPROCERS("PRO_CODE") = "F993" Then
      BIOCHEMISTRY1.chk_magn.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F865" Then
    BIOCHEMISTRY1.chk_livfuntest.Value = 1
    GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F866" Then
      BIOCHEMISTRY1.chk_prot.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F867" Then
      BIOCHEMISTRY1.chk_albumin.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F868" Then
      BIOCHEMISTRY1.chk_glob.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F869" Then
      BIOCHEMISTRY1.chk_alkphos.Value = 1
      GoTo quit
End If

If SELDMULPROCERS("PRO_CODE") = "F870" Then
      BIOCHEMISTRY1.chk_totbil.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F871" Then
      BIOCHEMISTRY1.chk_dirbil.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F872" Then
      BIOCHEMISTRY1.chk_indirbil.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F873" Then
      BIOCHEMISTRY1.chk_sgot.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F874" Then
      BIOCHEMISTRY1.chk_sgpt.Value = 1
      GoTo quit
End If
'If SELDMULPROCERS("PRO_CODE") = "F873" Then
'    biochemistry.chk_carenzy.Value = 1
'
'End If
'If SELDMULPROCERS("PRO_CODE") = "F874" Then
'      biochemistry.chk_LDH.Value = 1
'End If
'If SELDMULPROCERS("PRO_CODE") = "F875" Then
'      biochemistry.chk_ldh1.Value = 1
'End If
'If SELDMULPROCERS("PRO_CODE") = "F876" Then
'      biochemistry.chk_CPK.Value = 1
'End If
'If SELDMULPROCERS("PRO_CODE") = "F877" Then
'      biochemistry.chk_CPKMB.Value = 1
'End If
If SELDMULPROCERS("PRO_CODE") = "F880" Then
    BIOCHEMISTRY1.chk_panenzy.Value = 1
    GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F881" Then
'If SELDMULPROCERS("PRO_CODE") = "F988" Then
      BIOCHEMISTRY1.chk_amylase.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F882" Then
      BIOCHEMISTRY1.chk_lipase.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F893" Then
      BIOCHEMISTRY1.chk_iron.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F894" Then
    BIOCHEMISTRY1.chk_tibc.Value = 1
    GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F895" Then
      BIOCHEMISTRY1.chk_acidphosp.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F896" Then
      BIOCHEMISTRY1.chk_prostacid.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F970" Then
      BIOCHEMISTRY1.CHK_CALCI.Value = 1
      GoTo quit
End If

If SELDMULPROCERS("PRO_CODE") = "F969" Then
      BIOCHEMISTRY1.CHK_PHOS.Value = 1
      GoTo quit
End If

If SELDMULPROCERS("PRO_CODE") = "F884" Then
     BIOCHEMISTRY1.chk_csfsug.Value = 1
     GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F885" Then
      BIOCHEMISTRY1.chk_csfprot.Value = 1
      GoTo quit
End If

If SELDMULPROCERS("PRO_CODE") = "F886" Then
      BIOCHEMISTRY1.chk_csfchlor.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F887" Then
      BIOCHEMISTRY1.chk_csfldh.Value = 1
      GoTo quit
End If

If SELDMULPROCERS("PRO_CODE") = "F897" Then
      BIOCHEMISTRY1.chk_urinvma.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F899" Then
      'BIOCHEMISTRY1.chk_urob.Value = 1
      'GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F898" Then
      BIOCHEMISTRY1.chk_urin5hiaa.Value = 1
      GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F900" Then
     'BIOCHEMISTRY1.chk_Porphyrin.Value = 1
     'GoTo quit
End If
'************************************************
'3

'4
'************************************************
'''' BODY FLUIDS
'SUGAR
If SELDMULPROCERS("PRO_CODE") = "F962" Then
    BIOCHEMISTRY1.CHK_SUGARBF.Value = 1
    GoTo quit
End If
'UREA
If SELDMULPROCERS("PRO_CODE") = "F934" Then
    BIOCHEMISTRY1.CHK_UREABF.Value = 1
    GoTo quit
End If
'URIC ACID
If SELDMULPROCERS("PRO_CODE") = "F935" Then
    BIOCHEMISTRY1.CHK_URICACIDBF.Value = 1
    GoTo quit
End If
'CREATININE
If SELDMULPROCERS("PRO_CODE") = "F936" Then
    BIOCHEMISTRY1.CHK_CREATININEBF.Value = 1
    GoTo quit
End If
'PROTIEN
If SELDMULPROCERS("PRO_CODE") = "F963" Then
    BIOCHEMISTRY1.CHK_PROBF.Value = 1
    GoTo quit
End If
'ALBUMIN
If SELDMULPROCERS("PRO_CODE") = "F964" Then
    BIOCHEMISTRY1.CHK_ALBBF.Value = 1
    GoTo quit
End If
'DIRECT BIL
If SELDMULPROCERS("PRO_CODE") = "F941" Then
    BIOCHEMISTRY1.CHK_DIRECTBILBF.Value = 1
    GoTo quit
End If
'INDIRECT ALBUMIN
If SELDMULPROCERS("PRO_CODE") = "F942" Then
    BIOCHEMISTRY1.CHK_IBBF.Value = 1
    GoTo quit
End If

' GLOBULINS
If SELDMULPROCERS("PRO_CODE") = "F965" Then
    BIOCHEMISTRY1.CHK_GLOBF.Value = 1
    GoTo quit
End If
'ALK. PHOSPHATE
If SELDMULPROCERS("PRO_CODE") = "F966" Then
    BIOCHEMISTRY1.CHK_ALKBF.Value = 1
    GoTo quit
End If
' TOTAL BILIRUBIN
If SELDMULPROCERS("PRO_CODE") = "F940" Then
    BIOCHEMISTRY1.CHK_TOTALBILBF.Value = 1
    GoTo quit
End If
' SODIUM
If SELDMULPROCERS("PRO_CODE") = "F937" Then
    BIOCHEMISTRY1.CHK_NABF.Value = 1
    GoTo quit
End If
'CHLORIDE
If SELDMULPROCERS("PRO_CODE") = "F939" Then
    BIOCHEMISTRY1.CHK_CLBF.Value = 1
    GoTo quit
End If
'POATASSIUM
If SELDMULPROCERS("PRO_CODE") = "F938" Then
    BIOCHEMISTRY1.CHK_KBF.Value = 1
    GoTo quit
End If
'CHOLESTROL
If SELDMULPROCERS("PRO_CODE") = "F943" Then
    BIOCHEMISTRY1.CHK_CHOLBF.Value = 1
    GoTo quit
End If
' HDL CHOLESTEROL
If SELDMULPROCERS("PRO_CODE") = "F945" Then
    BIOCHEMISTRY1.CHK_HDLBF.Value = 1
    GoTo quit
End If
If SELDMULPROCERS("PRO_CODE") = "F876" Then
    BIOCHEMISTRY1.chk_ldh.Value = 1
    GoTo quit
End If

' TRIGLYCERIDES
If SELDMULPROCERS("PRO_CODE") = "F944" Then
    BIOCHEMISTRY1.CHK_TRIBF.Value = 1
    GoTo quit
End If
' LDL CHOLEST
If SELDMULPROCERS("PRO_CODE") = "F946" Then
    BIOCHEMISTRY1.CHK_LDLBF.Value = 1
    GoTo quit
End If
' ALYM
If SELDMULPROCERS("PRO_CODE") = "F881" Then
'If SELDMULPROCERS("PRO_CODE") = "F988" Then
    BIOCHEMISTRY1.CHK_ALBBF.Value = 1
    GoTo quit
End If
'LIPASE
If SELDMULPROCERS("PRO_CODE") = "F882" Then
    BIOCHEMISTRY1.CHK_LIPBF.Value = 1
    GoTo quit
End If
'LDH
If SELDMULPROCERS("PRO_CODE") = "F973" Then
    BIOCHEMISTRY1.CHK_LDHBF.Value = 1
    GoTo quit
End If
'SGOT
If SELDMULPROCERS("PRO_CODE") = "F967" Then
    BIOCHEMISTRY1.CHK_SGOTBF.Value = 1
    GoTo quit
End If
'SGPT
If SELDMULPROCERS("PRO_CODE") = "F968" Then
    BIOCHEMISTRY1.CHK_SGPTBF.Value = 1
  GoTo quit
End If

'''' BODY FLUIDS   '4
'*************************************************
''''STONE ANALYSIS  '5
If SELDMULPROCERS("PRO_CODE") = "F947" Then
    BIOCHEMISTRY1.chk_stoneanalysis.Value = 1
    GoTo quit
End If
''''STONE ANALYSIS   '5

''GTT
If SELDMULPROCERS("PRO_CODE") = "F844" Then
    BIOCHEMISTRY1.chk_gtt.Value = 1
    GoTo quit
End If
'GTT
quit:
     
End Sub

Public Sub staffprint()
        If viewflag = True Then
            'HT
            If txt_modcode = "HT" Then
                If SELDREQUISITNFLAG = False Then
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                    SELDREQUISITNFLAG = False
                Else
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    seldrequisitn_rs.MoreResults
                    seldrequisitn_rs.Requery
                End If
                If SELTA101FALG = False Then
                    'selta101
                    SELTA101_QR(0) = Msk_caseno.FormattedText
                    Set SELTA101_RS = SELTA101_QR.OpenResultset(2, 3)
                    SELTA101FALG = True
                Else
                    'selta101
                    SELTA101_QR(0) = Msk_caseno.FormattedText
                    Set SELTA101_RS = SELTA101_QR.OpenResultset(2, 3)
                    SELTA101_RS.MoreResults
                    SELTA101_RS.Requery
                End If
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPHTSTAFF.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & cmb_reqno.Text & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                'If Len(seldrequisitn_rs("LABNO")) > 0 Then
                If Trim(seldrequisitn_rs("LABNO")) = "yes" Then
                    CrystalReport1.Formulas(6) = "NAME = '" & Trim(seldrequisitn_rs("NMETERIALS")) & "'"
                    CrystalReport1.Formulas(7) = "RELATION = 'Relation' "
                Else
                    CrystalReport1.Formulas(6) = "NAME = '" & Trim(txt_name.Text) & "'"
                    CrystalReport1.Formulas(7) = "RELATION ='" & " " & "'"
                End If
                CrystalReport1.Formulas(8) = "ENTERBY='" & Trim(enter_by1) & "'"
                
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            End If
            '*** HLA TYPING
            '*** DTM LAB
            
            
            'SS
            If txt_modcode = "SS" Then
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPSSstaff.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & cmb_reqno.Text & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & TEXT1.Text & "'" ' txt_pathno & "'"


'                If Trim(prepath) = "" Then
'                    CrystalReport1.Formulas(5) = "PATH = '" & "" & " '"
'                Else
'                    CrystalReport1.Formulas(5) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
'                End If
                If txt_modcode.Text = "SS" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR SCREENING SERVICES" & " '"
                End If
                CrystalReport1.Formulas(7) = "NAME = '" & Trim(txt_name.Text) & " '"
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            End If
            '************************ED
                If txt_modcode = "ED" Then
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPEDstaff.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & cmb_reqno.Text & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
               
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_sex & "/" & txt_Age & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PROV = '" & txt_indctinvstg & "'"
                CrystalReport1.Formulas(7) = "ENTERBY='" & Trim(enter_by1) & "'"
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
             End If
            '*** ENDOSCOPY

                
            '************************ED
            
            
             ''**** BLOOD PRODUCT
            If txt_modcode = "BP" Then
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPBPSTAFF.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & cmb_reqno.Text & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno.Text & "'"
                
                
'                If Trim(prepath) = "" Then
'                    CrystalReport1.Formulas(5) = "PATH = '" & "" & " '"
'                Else
'                    CrystalReport1.Formulas(5) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
'                End If
                CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR BLOOD PRODUCTS" & " '"
                CrystalReport1.Formulas(7) = "NAME = '" & Trim(txt_name.Text) & " '"
                CrystalReport1.Formulas(8) = "PROV = '" & txt_indctinvstg & "'"
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            End If
            
            ''ST REPORT
            If txt_modcode = "ST" Or txt_modcode = "2D" Or txt_modcode = "PF" Or txt_modcode = "EK" Then    'SUN
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPECST.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & cmb_reqno.Text & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                'CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                
                
                If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(5) = "PATH = '" & "" & " '"
                Else
                    CrystalReport1.Formulas(5) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
                End If
'                If txt_modcode.Text = "EK" Then
'                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR ECG" & " '"
'                End If
                If txt_modcode.Text = "2D" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR 2DECHO" & " '"
                End If
                If txt_modcode.Text = "ST" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR STRESS TEST " & " '"
                End If
                CrystalReport1.Formulas(7) = "NAME = '" & txt_name.Text & " '"
'                If txt_modcode.Text = "PF" Then
'                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR PFT TEST " & " '"
'                End If
                 CrystalReport1.Formulas(8) = "ENTERBY='" & Trim(enter_by1) & "'"
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            End If
            
            If txt_modcode = "SP" Then
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPS.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & cmb_reqno.Text & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_sex & "/" & txt_Age & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                I = InStr(cmb_sites, "$")
                If I = 0 Then
                    CrystalReport1.Formulas(6) = "SITE = '" & cmb_sites & "'"
                Else
                    CrystalReport1.Formulas(6) = "SITE = '" & Mid(cmb_sites, 1, I - 1) & "'"
                End If
                CrystalReport1.Formulas(7) = "LATER = '" & cmb_laterality1 & "'"
                CrystalReport1.Formulas(8) = "OPPROC = '" & TXT_OPPROCEDURE & "'"
                CrystalReport1.Formulas(9) = "METAS = '" & txt_metastasis & "'"
                CrystalReport1.Formulas(10) = "LATER1 = '" & cmb_laterality & "'"
                CrystalReport1.Formulas(11) = "OUTSP = '" & txt_outspecimen & "'"
                CrystalReport1.Formulas(12) = "OUTBLK = '" & cmb_outblock & "'"
                CrystalReport1.Formulas(13) = "OUTSLID = '" & cmb_outslide & "'"
                CrystalReport1.Formulas(14) = "USRID = '" & str & "'"
                CrystalReport1.Formulas(15) = "PROV = '" & txt_indctinvstg & "'"
                If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(16) = "PATH = '" & "" & " '"
                Else
                    CrystalReport1.Formulas(16) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
                End If
                CrystalReport1.Formulas(17) = "ENTERBY='" & Trim(enter_by1) & "'"
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            End If
            
            
            
            If txt_modcode = "IH" Then
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPIHs.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & cmb_reqno.Text & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In Patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_sex & "/" & txt_Age & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                CrystalReport1.Formulas(6) = "OUTBLK = '" & cmb_Ihsoutsideblock & "'"
                CrystalReport1.Formulas(7) = "OUTSLID = '" & cmb_Ihsoutslide & "'"
                 If seldihreportflag = False Then
                    seldihreport_qr(0) = req_chglen(cmb_reqno)
                    Set seldihreport_rs = seldihreport_qr.OpenResultset(1, 3)
                    seldihreportflag = True
                 Else
                    seldihreport_qr(0) = req_chglen(cmb_reqno.Text)
                    seldihreport_rs.MoreResults
                    seldihreport_rs.Requery
                 End If
                 If seldihreport_rs.RowCount > 0 Then
                 str1 = ""
                    While Not seldihreport_rs.EOF
                        str1 = str1 & Trim(seldihreport_rs("testorder")) & ", "
                        seldihreport_rs.MoveNext
                    Wend
                 End If
                I = Len(Trim(str1))
                If I = 0 Then
                    CrystalReport1.Formulas(8) = "TEST = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(8) = "TEST = '" & Left(Trim(str1), I - 1) & "'"
                End If
                CrystalReport1.Formulas(9) = "USRID = '" & str & "'"
                CrystalReport1.Formulas(10) = "PROV = '" & txt_indctinvstg & "'"
                If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(11) = "PATH = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(11) = "PATH = '" & "Previous Path No." & Left(Trim(prepath), 200) & "'"
                End If
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            End If
            If txt_modcode.Text = "FA" Then
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPFAS.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & cmb_reqno.Text & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_sex & "/" & txt_Age & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                I = InStr(cmb_sites, "$")
                If I = 0 Then
                    CrystalReport1.Formulas(6) = "SITE = '" & cmb_fnsite & "'"
                Else
                    CrystalReport1.Formulas(6) = "SITE = '" & Mid(cmb_fnsite, 1, I - 1) & "'"
                End If
                CrystalReport1.Formulas(7) = "LATER = '" & cmb_fnlaterality & "'"
                CrystalReport1.Formulas(8) = "METAS = '" & txt_fnmetastasis & "'"
                CrystalReport1.Formulas(9) = "OPPROC = '" & cmb_fnac & "'"
                CrystalReport1.Formulas(10) = "LATER1 = '" & cmb_fnlaterality1 & "'"
                CrystalReport1.Formulas(11) = "OUTSP = '" & TXT_FNSLIDEDRY & "'"
                CrystalReport1.Formulas(12) = "OUTBLK = '" & TXT_FNOSLIDEWET & "'"
                CrystalReport1.Formulas(13) = "OUTSLID = '" & cmb_outslide & "'"
                CrystalReport1.Formulas(14) = "USRID = '" & str & "'"
                CrystalReport1.Formulas(15) = "PROV = '" & txt_indctinvstg & "'"
                If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(16) = "PATH = '" & "" & " '"
                Else
                    CrystalReport1.Formulas(16) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
                End If
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            End If
            If txt_modcode = "CG" Or txt_modcode = "CN" Then
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPCGS.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & cmb_reqno.Text & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_sex & "/" & txt_Age & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                I = InStr(cmb_cnsite, "$")
                If I = 0 Then
                    CrystalReport1.Formulas(6) = "SITE = '" & cmb_cnsite & "'"
                Else
                    CrystalReport1.Formulas(6) = "SITE = '" & Mid(cmb_cnsite, 1, I - 1) & "'"
                End If
                CrystalReport1.Formulas(7) = "LATER = '" & cmb_claterality & "'"
                CrystalReport1.Formulas(8) = "OPPROC = '" & txt_trtdetail & "'"
                CrystalReport1.Formulas(9) = "METAS = '" & txt_metastasis & "'"
                CrystalReport1.Formulas(10) = "LATER1 = '" & txt_scopyfind & "'"
                CrystalReport1.Formulas(11) = "OUTSP = '" & cmb_cyoutspecimen & "'"
                CrystalReport1.Formulas(12) = "OUTBLK = '" & cmb_cnoslides & "'"
                CrystalReport1.Formulas(13) = "OUTSLID = '" & cmb_coutslides & "'"
                CrystalReport1.Formulas(14) = "USRID = '" & str & "'"
                CrystalReport1.Formulas(15) = "PROV = '" & txt_indctinvstg & "'"
                If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(16) = "PATH = '" & "" & " '"
                Else
                    CrystalReport1.Formulas(16) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
                End If
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            End If
            '*** BIOCHEMISTRY
            If txt_modcode = "BC" Or txt_modcode = "EL" Or txt_modcode = "TM" Then
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPBCC.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & cmb_reqno.Text & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
'                If txt_status = "ST" Then
'                    CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
'                End If
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_sex & "/" & txt_Age & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                
                
'                If Trim(prepath) = "" Then
'                    CrystalReport1.Formulas(5) = "PATH = '" & "" & " '"
'                Else
'                    CrystalReport1.Formulas(5) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
'                End If
                If txt_modcode.Text = "BC" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR BIOCHEMSITRY" & " '"
                End If
                If txt_modcode.Text = "EL" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR EMERGENCY LAB" & " '"
                End If
                If txt_modcode.Text = "TM" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR TUMOUR MARKER " & " '"
                End If
                CrystalReport1.Formulas(7) = "ENTERBY='" & Trim(enter_by1) & "'"
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            End If
            '*** BIOCHEMSITRY
            
            '************** BODY  FLUIDS
                If txt_modcode = "BF" Then
                    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPBFNSTAFF.RPT"
                    CrystalReport1.Formulas(0) = "REQNO = '" & cmb_reqno.Text & "'"
                    CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                    If txt_status = "OP" Then
                        CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                    Else
                        CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                    End If
                    '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                    '*****
                    CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                    If Trim(txt_status.Text) = "IP" Then
                        CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                    Else
                        CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                    End If
                    CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                
                    If txt_modcode.Text = "BF" Then
                      CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR BODY FLUIDS " & " '"
                    End If
                    CrystalReport1.Formulas(7) = "ENTERBY='" & Trim(enter_by1) & "'"
                    CrystalReport1.WindowState = crptMaximized
                    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                    CrystalReport1.Action = 1
                End If
            
            '************** BODY FLUIDS
            
            '************* BONE MARROW ASPIRATION REPORT
                '*** bone marrow aspiration
            If txt_modcode = "BM" Then
                If SELDBMACIND_FLAG = False Then
                    SELDBMACIND_QR(0) = req_chglen(cmb_reqno.Text)
                    Set SELDBMACIND_RS = SELDBMACIND_QR.OpenResultset(1, 3)
                    SELDBMACIND_FLAG = True
                Else
                    SELDBMACIND_QR(0) = req_chglen(cmb_reqno.Text)
                    SELDBMACIND_RS.MoreResults
                    SELDBMACIND_RS.Requery
                End If
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
               ' CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPBMN.RPT"
               CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPHLSTAFF.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & cmb_reqno.Text & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                
                If Not SELDBMACIND_RS.EOF Then
                    If SELDBMACIND_RS("NEWCASE") <> "" Then
                        CrystalReport1.Formulas(7) = "NEWCASE = '" & SELDBMACIND_RS("NEWCASE") & "'"
                    End If
                    If SELDBMACIND_RS("POSTTRANS") <> "" Then
                        CrystalReport1.Formulas(8) = "POSTTRANS = '" & SELDBMACIND_RS("POSTTRANS") & "'"
                    End If
                    If SELDBMACIND_RS("DAY10MARR") <> "" Then
                        CrystalReport1.Formulas(9) = "DAY10MARR = '" & SELDBMACIND_RS("DAY10MARR") & "'"
                    End If
                    If SELDBMACIND_RS("PARTREATED") <> "" Then
                        CrystalReport1.Formulas(10) = "PARTREATED = '" & SELDBMACIND_RS("PARTREATED") & "'"
                    End If
                    If SELDBMACIND_RS("TSTEROIDS") <> "" Then
                        CrystalReport1.Formulas(11) = "TSTEROIDS = '" & SELDBMACIND_RS("TSTEROIDS") & "'"
                    End If
                    If SELDBMACIND_RS("RELAPSE") <> "" Then
                        CrystalReport1.Formulas(12) = "RELAPSE = '" & SELDBMACIND_RS("RELAPSE") & "'"
                    End If
                    If SELDBMACIND_RS("PINDUCTION") <> "" Then
                        CrystalReport1.Formulas(13) = "PINDUCTION = '" & SELDBMACIND_RS("PINDUCTION") & "'"
                    End If
                    If SELDBMACIND_RS("POSTBMT") <> "" Then
                        CrystalReport1.Formulas(14) = "POSTBMT = '" & SELDBMACIND_RS("POSTBMT") & "'"
                    End If
                    If SELDBMACIND_RS("OTHER") <> "" Then
                        CrystalReport1.Formulas(15) = "OTHER = 'Other : " & SELDBMACIND_RS("OTHER") & "'"
                    End If
                End If
                CrystalReport1.Formulas(16) = "PROV = '" & txt_indctinvstg & "'"
'                If Trim(prepath) = "" Then
'                    CrystalReport1.Formulas(5) = "PATH = '" & "" & " '"
'                Else
'                    CrystalReport1.Formulas(5) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
'                End If
                If txt_modcode.Text = "BM" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR BONE MARROW ASPIRATION " & " '"
                End If
                CrystalReport1.Formulas(17) = "ENTERBY='" & Trim(enter_by1) & "'"
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            End If
            
            
            
            
            '*************BONE MARROW ASPITRATION REPORT
            
                       
            '***** HL HEMOTOLOGY REPORT
                If txt_modcode = "HL" Then
                    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPHLNSTAFF.RPT"
                    CrystalReport1.Formulas(0) = "REQNO = '" & cmb_reqno.Text & "'"
                    CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                    If txt_status = "OP" Then
                        CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                    Else
                        CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                    End If
                    '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                    '*****
                    CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                    If Trim(txt_status.Text) = "IP" Then
                        CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                    Else
                        CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                    End If
                    CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                              
'                   If Trim(prepath) = "" Then
'                        CrystalReport1.Formulas(5) = "PATH = '" & "" & " '"
'                   Else
'                       CrystalReport1.Formulas(5) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
'                   End If
                    If txt_modcode.Text = "HL" Then
                        CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR HAEMATOLOGY" & " '"
                    End If
                    CrystalReport1.Formulas(7) = "ENTERBY='" & Trim(enter_by1) & "'"
                    CrystalReport1.WindowState = crptMaximized
                    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                    CrystalReport1.Action = 1
                End If
            '***** HL HEMOTOLOGY REPORT
            
            '*** MICROBIOLOGY
            If txt_modcode = "BA" Or txt_modcode = "MB" Or txt_modcode = "SE" Or txt_modcode = "MC" Or txt_modcode = "CM" Or txt_modcode = "MO" Or txt_modcode = "XX" Then
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPMBs.RPT"   'sun
                CrystalReport1.Formulas(0) = "REQNO = '" & cmb_reqno.Text & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_sex & "/" & txt_Age & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                
                
                If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(5) = "PATH = '" & "" & " '"
                Else
                    CrystalReport1.Formulas(5) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
                End If
                If txt_modcode.Text = "BA" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR BACTERIOLOGY" & " '"
                End If
                If txt_modcode.Text = "MB" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR MYCOBACTERIOLOGY" & " '"
                End If
                If txt_modcode.Text = "SE" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR SEROLOGY " & " '"
                End If
                If txt_modcode.Text = "MC" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR MYCOLOGY " & " '"
                End If
                If txt_modcode.Text = "CM" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR CLINICAL MICROBIOLOGY" & " '"
                End If
                If txt_modcode.Text = "MO" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR MOLECULAR MICROBIOLOGY " & " '"
                End If
                If txt_modcode.Text = "XX" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR OTHERS " & " '"
                End If
                   CrystalReport1.Formulas(7) = "icu = '" & Trim(MCICU) & "'"
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            End If
            '*** MICROBIOLOGY
        Else
            Data1.Recordset.MoveFirst
            While Not Data1.Recordset.EOF
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(Data1.Recordset.Fields("REQNO")) & "'"
                '****HT Report
                If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "HT" Then
                    If SELDREQUISITNFLAG = False Then
                        seldrequisitn_qr(0) = req_chglen(Trim(Data1.Recordset.Fields("REQNO")))
                        Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                        SELDREQUISITNFLAG = False
                    Else
                        seldrequisitn_qr(0) = req_chglen(Trim(Data1.Recordset.Fields("REQNO")))
                        seldrequisitn_rs.MoreResults
                        seldrequisitn_rs.Requery
                    End If
                    If SELTA101FALG = False Then
                        'selta101
                        SELTA101_QR(0) = Msk_caseno.FormattedText
                        Set SELTA101_RS = SELTA101_QR.OpenResultset(2, 3)
                        SELTA101FALG = True
                    Else
                        'selta101
                        SELTA101_QR(0) = Msk_caseno.FormattedText
                        Set SELTA101_RS = SELTA101_QR.OpenResultset(2, 3)
                        SELTA101_RS.MoreResults
                        SELTA101_RS.Requery
                    End If
                    'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPHTSTAFF.RPT"
                    CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
                    CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                    If txt_status = "OP" Then
                        CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                    Else
                        CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                    End If
                    '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                    '*****
                    
                    CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                    If Trim(txt_status.Text) = "IP" Then
                        CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                    Else
                        CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                    End If
                    CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                    'If Len(seldrequisitn_rs("LABNO")) > 0 Then
                    If Trim(seldrequisitn_rs("LABNO")) = "yes" Then
                        CrystalReport1.Formulas(6) = "NAME = '" & Trim(seldrequisitn_rs("NMETERIALS")) & "'"
                        CrystalReport1.Formulas(7) = "RELATION = 'Relation' "
                    Else
                        CrystalReport1.Formulas(6) = "NAME = '" & Trim(txt_name.Text) & "'"
                        CrystalReport1.Formulas(7) = "RELATION ='" & " " & "'"
                    End If
                    'CrystalReport1.Action = 1
                End If
                '''*********ST
                If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "ST" Or Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "2D" Or Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "PF" Or Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "EK" Then
                'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & Data1.Recordset.Fields("REQNO") & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPECST.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                'CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                
                
                If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(5) = "PATH = '" & "" & " '"
                Else
                    CrystalReport1.Formulas(5) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
                End If
                If txt_modcode.Text = "EK" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR ECG" & " '"
                End If
                If txt_modcode.Text = "2D" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR 2DECHO" & " '"
                End If
                If txt_modcode.Text = "ST" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR STRESS TEST " & " '"
                End If
                CrystalReport1.Formulas(7) = "NAME = '" & txt_name.Text & " '"
'                If txt_modcode.Text = "PF" Then
'                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR PFT TEST " & " '"
'                End If
                'CrystalReport1.Action = 1
            End If
                
                
                '**** SP REPORT
                If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "SP" Then
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPSPNS.RPT"
                    CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
'                    I = InStr(cmb_modality.Text, "$")
'                    CRYSTALREPORT1.Formulas(1) = "MOD = '" & Left(cmb_modality.Text, I - 1) & "'"
                    If txt_status = "OP" Then
                        CrystalReport1.Formulas(1) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                    Else
                        CrystalReport1.Formulas(1) = "ST = '" & txt_catgory & "/" & " In Patient" & "'"
                    End If
                    '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                    '*****
                    CrystalReport1.Formulas(2) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                    If Trim(txt_status.Text) = "IP" Then
                        CrystalReport1.Formulas(3) = "WRBED = '" & txt_sex & "/" & txt_Age & "'"
                    Else
                        CrystalReport1.Formulas(3) = "WRBED = '" & " " & "'"
                    End If
                    CrystalReport1.Formulas(4) = "USRID = '" & str & "'"
                    CrystalReport1.Formulas(5) = "PROV = '" & txt_indctinvstg & "'"
                    CrystalReport1.Formulas(6) = "PATHNO = '" & txt_sppathno & "'"
                    I = InStr(cmb_sites, "$")
                    If I = 0 Then
                     If Len(Trim(cmb_sites)) > 0 Then
                        CrystalReport1.Formulas(7) = "SITE = '" & UCase(Trim(cmb_sites.Text)) & "'"
                     Else
                        CrystalReport1.Formulas(7) = "SITE = '" & "" & "'"
                     End If
                    Else
                     CrystalReport1.Formulas(7) = "SITE = '" & Mid(cmb_sites, 1, I - 1) & "'"
                    End If
                    If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(8) = "PATH = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(8) = "PATH = '" & "Previous Path No." & Left(Trim(prepath), 200) & "'"
                End If
               End If
                '**** SP REPORT
                 '**** IH REPORT
                 If Mid(Data1.Recordset.Fields("REQNO"), 5, 2) = "IH" Then
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPIHNS.RPT"
                    CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
                    I = InStr(cmb_modality.Text, "$")
                    CrystalReport1.Formulas(1) = "MOD = '" & Left(cmb_modality.Text, I - 1) & "'"
                    If txt_status = "OP" Then
                        CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                    Else
                        CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & "In Patient" & "'"
                    End If
                    '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                    '*****
                    CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                    If Trim(txt_status.Text) = "IP" Then
                        CrystalReport1.Formulas(4) = "WRBED = '" & txt_sex & "/" & txt_Age & "'"
                    Else
                        CrystalReport1.Formulas(5) = "WRBED = '" & " " & "'"
                    End If
                     CrystalReport1.Formulas(6) = "outblock = '" & cmb_Ihsoutsideblock & "'"
                     CrystalReport1.Formulas(7) = "outslide = '" & cmb_Ihsoutslide & "'"

                     If seldihreportflag = False Then
                    seldihreport_qr(0) = req_chglen(txt_ihreqno)
                    Set seldihreport_rs = seldihreport_qr.OpenResultset(1, 3)
                    seldihreportflag = True
                 Else
                    seldihreport_qr(0) = req_chglen(cmb_reqno.Text)
                    seldihreport_rs.MoreResults
                    seldihreport_rs.Requery
                 End If
                 If seldihreport_rs.RowCount > 0 Then
                 str1 = ""
                    While Not seldihreport_rs.EOF
                        str1 = str1 & Trim(seldihreport_rs("testorder")) & ", "
                        seldihreport_rs.MoveNext
                    Wend
                 End If
                    I = Len(Trim(str1))
                    If I = 0 Then
                        CrystalReport1.Formulas(8) = "TEST = '" & "" & "'"
                    Else
                        CrystalReport1.Formulas(8) = "TEST = '" & Left(Trim(str1), I - 1) & "'"
                    End If
                    CrystalReport1.Formulas(9) = "USRID = '" & str & "'"
                    CrystalReport1.Formulas(10) = "PROV = '" & txt_indctinvstg & "'"
                    CrystalReport1.Formulas(11) = "PATHNO = '" & txt_ihpathno & "'"
                    If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(12) = "PATH = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(12) = "PATH = '" & "Previous Path No." & Left(Trim(prepath), 200) & "'"
                End If
                 End If
                 '**** IH REPORT
                 '**** FA REPORT
                 If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "FA" Then
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPFANS.RPT"
                    CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
'                    I = InStr(cmb_modality.Text, "$")
'                    CRYSTALREPORT1.Formulas(1) = "MOD = '" & Left(cmb_modality.Text, I - 1) & "'"
                    If txt_status = "OP" Then
                        CrystalReport1.Formulas(1) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                    Else
                        CrystalReport1.Formulas(1) = "ST = '" & txt_catgory & "/" & " In Patient" & "'"
                    End If
                    '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                    '*****
                    CrystalReport1.Formulas(2) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                    If Trim(txt_status.Text) = "IP" Then
                        CrystalReport1.Formulas(3) = "WRBED = '" & txt_sex & "/" & txt_Age & "'"
                    Else
                        CrystalReport1.Formulas(3) = "WRBED = '" & " " & "'"
                    End If
                    CrystalReport1.Formulas(4) = "USRID = '" & str & "'"
                    CrystalReport1.Formulas(5) = "PROV = '" & txt_indctinvstg & "'"
                    CrystalReport1.Formulas(6) = "PATHNO = '" & txt_sppathno & "'"
                    I = InStr(cmb_fnsite, "$")
                    If I = 0 Then
                     If Len(Trim(cmb_fnsite)) > 0 Then
                        CrystalReport1.Formulas(7) = "SITE = '" & UCase(Trim(cmb_fnsite.Text)) & "'"
                     Else
                        CrystalReport1.Formulas(7) = "SITE = '" & "" & "'"
                     End If
                    Else
                     CrystalReport1.Formulas(7) = "SITE = '" & Mid(cmb_fnsite, 1, I - 1) & "'"
                    End If
                    If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(8) = "PATH = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(8) = "PATH = '" & "Previous Path No." & Left(Trim(prepath), 200) & "'"
                End If
               End If
                 '**** FA REPORT
                 '**** CG OR CN REPORT
                 If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "CG" Or Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "CN" Then
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPCGNS.RPT"
                    CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
'                    I = InStr(cmb_modality.Text, "$")
'                    CRYSTALREPORT1.Formulas(1) = "MOD = '" & Left(cmb_modality.Text, I - 1) & "'"
                    If txt_status = "OP" Then
                        CrystalReport1.Formulas(1) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                    Else
                        CrystalReport1.Formulas(1) = "ST = '" & txt_catgory & "/" & " In Patient" & "'"
                    End If
                    '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                    '*****
                    CrystalReport1.Formulas(2) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                    If Trim(txt_status.Text) = "IP" Then
                        CrystalReport1.Formulas(3) = "WRBED = '" & txt_sex & "/" & txt_Age & "'"
                    Else
                        CrystalReport1.Formulas(3) = "WRBED = '" & " " & "'"
                    End If
                    CrystalReport1.Formulas(4) = "USRID = '" & str & "'"
                    CrystalReport1.Formulas(5) = "PROV = '" & txt_indctinvstg & "'"
                    CrystalReport1.Formulas(6) = "PATHNO = '" & txt_sppathno & "'"
                    I = InStr(cmb_cnsite, "$")
                    If I = 0 Then
                     If Len(Trim(cmb_cnsite)) > 0 Then
                        CrystalReport1.Formulas(7) = "SITE = '" & UCase(Trim(cmb_cnsite.Text)) & "'"
                     Else
                        CrystalReport1.Formulas(7) = "SITE = '" & "" & "'"
                     End If
                    Else
                     CrystalReport1.Formulas(7) = "SITE = '" & Mid(cmb_cnsite, 1, I - 1) & "'"
                    End If
                    If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(8) = "PATH = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(8) = "PATH = '" & "Previous Path No." & Left(Trim(prepath), 200) & "'"
                End If
               End If
               
                 '**** CG OR CN REPORT
                 '**** BC OR EL OR TM REPORT
                If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "BC" Or Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "EL" Or Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "TM" Then
                'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(Data1.Recordset.Fields("REQNO")) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPBCC.RPT"
                'CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
                CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
                    I = InStr(cmb_modality.Text, "$")
                    CrystalReport1.Formulas(1) = "MOD = '" & Left(cmb_modality.Text, I - 1) & "'"
                    If txt_status = "OP" Then
                        CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                    Else
                        CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & "In Patient" & "'"
                    End If
                    '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                    '*****
                    CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                    If Trim(txt_status.Text) = "IP" Then
                        CrystalReport1.Formulas(4) = "WRBED = '" & txt_sex & "/" & txt_Age & "'"
                    Else
                        CrystalReport1.Formulas(5) = "WRBED = '" & " " & "'"
                    End If
                
                
                
'                If Trim(prepath) = "" Then
'                    CrystalReport1.Formulas(6) = "PATH = '" & "" & " '"
'                Else
'                    CrystalReport1.Formulas(6) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
'                End If
                If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "BC" Then
                   CrystalReport1.Formulas(7) = "CAP = '" & "REQUISITION FOR BIOCHEMSITRY" & " '"
                End If
                If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "EL" Then
                   CrystalReport1.Formulas(7) = "CAP = '" & "REQUISITION FOR EMERGENCY LAB" & " '"
                End If
                If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "TM" Then
                   CrystalReport1.Formulas(7) = "CAP = '" & "REQUISITION FOR TUMOUR MARKER " & " '"
                End If
                CrystalReport1.Formulas(8) = "ENTERBY='" & Trim(enter_by1) & "'"
            End If
                 '**** BC OR EL OR TM REPORT
                 '**** MICROBIOLOGY
                 If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "BA" Or Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "MB" Or Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "SE" Or Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "MC" Or Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "CM" Or Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "MO" Or Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "XX" Then
                
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPMBS.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
                I = InStr(cmb_modality.Text, "$")
               '     CrystalReport1.Formulas(1) = "MOD = '" & Left(Cmb_modality.Text, i - 1) & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                    '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_sex & "/" & txt_Age & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                
                
                If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(5) = "PATH = '" & "" & " '"
                Else
                    CrystalReport1.Formulas(5) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
                End If
                If txt_modcode.Text = "BA" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR BACTERIOLOGY" & " '"
                End If
                If txt_modcode.Text = "MB" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR MYCOBACTERIOLOGY" & " '"
                End If
                If txt_modcode.Text = "SE" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR SEROLOGY " & " '"
                End If
                If txt_modcode.Text = "MC" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR MYCOLOGY " & " '"
                End If
                If txt_modcode.Text = "CM" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR CLINICAL MICROBIOLOGY" & " '"
                End If
                If txt_modcode.Text = "MO" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR MOLECULAR MICROBIOLOGY " & " '"
                End If
                If txt_modcode.Text = "XX" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR OTHERS " & " '"
                End If
                   CrystalReport1.Formulas(7) = "icu = '" & Trim(MCICU) & "'"
            End If
                 '**** SS
            If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "SS" Then
                 'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(Data1.Recordset.Fields("REQNO")) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPSSstaff.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                    '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & TEXT1.Text & "'" ' txt_pathno & "'"


'                If Trim(prepath) = "" Then
'                    CrystalReport1.Formulas(5) = "PATH = '" & "" & " '"
'                Else
'                    CrystalReport1.Formulas(5) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
'                End If
                If txt_modcode.Text = "SS" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR SCREENING SERVICES" & " '"
                End If
                CrystalReport1.Formulas(7) = "NAME = '" & Trim(txt_name.Text) & " '"
            End If
            
            ' ************* BODY FLUIDS
                If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "BF" Then
                'CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPBFNSTAFF.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                    '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                
                
'                If Trim(prepath) = "" Then
'                    CrystalReport1.Formulas(5) = "PATH = '" & "" & " '"
'                Else
'                    CrystalReport1.Formulas(5) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
'                End If
                If txt_modcode.Text = "BF" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR BODY FLUIDS " & " '"
                End If
                CrystalReport1.Formulas(7) = "ENTERBY='" & Trim(enter_by1) & "'"
                'CRYSTALREPORT1.WindowState = crptMaximized
                'CRYSTALREPORT1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                'CRYSTALREPORT1.Action = 1
            End If
            
            ' ************* BODY FLUIDS
            
            '*************** BONE MARROW STAFF REPORT
                '*** bone marrow aspiration
            If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "BM" Then
                If SELDBMACIND_FLAG = False Then
                    SELDBMACIND_QR(0) = req_chglen(cmb_reqno.Text)
                    Set SELDBMACIND_RS = SELDBMACIND_QR.OpenResultset(1, 3)
                    SELDBMACIND_FLAG = True
                Else
                    SELDBMACIND_QR(0) = req_chglen(cmb_reqno.Text)
                    SELDBMACIND_RS.MoreResults
                    SELDBMACIND_RS.Requery
                End If
                'CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                'CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPBMN.RPT"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPHLSTAFF.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                    '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                
                If Not SELDBMACIND_RS.EOF Then
                    If SELDBMACIND_RS("NEWCASE") <> "" Then
                        CrystalReport1.Formulas(7) = "NEWCASE = '" & SELDBMACIND_RS("NEWCASE") & "'"
                    End If
                    If SELDBMACIND_RS("POSTTRANS") <> "" Then
                        CrystalReport1.Formulas(8) = "POSTTRANS = '" & SELDBMACIND_RS("POSTTRANS") & "'"
                    End If
                    If SELDBMACIND_RS("DAY10MARR") <> "" Then
                        CrystalReport1.Formulas(9) = "DAY10MARR = '" & SELDBMACIND_RS("DAY10MARR") & "'"
                    End If
                    If SELDBMACIND_RS("PARTREATED") <> "" Then
                        CrystalReport1.Formulas(10) = "PARTREATED = '" & SELDBMACIND_RS("PARTREATED") & "'"
                    End If
                    If SELDBMACIND_RS("TSTEROIDS") <> "" Then
                        CrystalReport1.Formulas(11) = "TSTEROIDS = '" & SELDBMACIND_RS("TSTEROIDS") & "'"
                    End If
                    If SELDBMACIND_RS("RELAPSE") <> "" Then
                        CrystalReport1.Formulas(12) = "RELAPSE = '" & SELDBMACIND_RS("RELAPSE") & "'"
                    End If
                    If SELDBMACIND_RS("PINDUCTION") <> "" Then
                        CrystalReport1.Formulas(13) = "PINDUCTION = '" & SELDBMACIND_RS("PINDUCTION") & "'"
                    End If
                    If SELDBMACIND_RS("POSTBMT") <> "" Then
                        CrystalReport1.Formulas(14) = "POSTBMT = '" & SELDBMACIND_RS("POSTBMT") & "'"
                    End If
                    If SELDBMACIND_RS("OTHER") <> "" Then
                        CrystalReport1.Formulas(15) = "OTHER = 'Other : " & SELDBMACIND_RS("OTHER") & "'"
                    End If
                End If
                CrystalReport1.Formulas(16) = "PROV = '" & txt_indctinvstg & "'"
'                If Trim(prepath) = "" Then
'                    CrystalReport1.Formulas(5) = "PATH = '" & "" & " '"
'                Else
'                    CrystalReport1.Formulas(5) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
'                End If
                If txt_modcode.Text = "BM" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR BONE MARROW ASPIRATION " & " '"
                End If
                CrystalReport1.Formulas(17) = "ENTERBY='" & Trim(enter_by1) & "'"
                'CRYSTALREPORT1.WindowState = crptMaximized
                'CRYSTALREPORT1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                'CRYSTALREPORT1.Action = 1
            End If
            
            
            '*************** BONE MARRROW STAFF REPORT
            '**************** ED
                If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "ED" Then
                'CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPEDSTAFF.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_sex & "/" & txt_Age & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PROV = '" & txt_indctinvstg & "'"
                CrystalReport1.Formulas(7) = "ENTERBY='" & Trim(enter_by1) & "'"
                'CRYSTALREPORT1.WindowState = crptMaximized
                'CRYSTALREPORT1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                'CRYSTALREPORT1.Action = 1
             End If
            '*** ENDOSCOPY
           '**************** ED
           '' HL STAFF REPORT
                If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "HL" Then
                'CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPHLNSTAFF.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno & "'"
                               
'                If Trim(prepath) = "" Then
'                    CrystalReport1.Formulas(5) = "PATH = '" & "" & " '"
'                Else
'                    CrystalReport1.Formulas(5) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
'                End If
                If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "HL" Then
                   CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR HAEMATOLOGY" & " '"
                End If
                CrystalReport1.Formulas(7) = "ENTERBY='" & Trim(enter_by1) & "'"
                'CrystalReport1.WindowState = crptMaximized
                'CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                'CrystalReport1.Action = 1
            End If
            
            
            
            '' HL STAFF REPORT
                      
            ''BP 'AM
            If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "BP" Then
                'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(cmb_reqno) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPBPSTAFF.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
                CrystalReport1.Formulas(1) = "MOD = '" & cmb_modality.Text & "'"
                If txt_status = "OP" Then
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                Else
                    CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " In patient" & "'"
                End If
                '********  on 31/12/04
                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If
                '*****
                CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                If Trim(txt_status.Text) = "IP" Then
                    CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                Else
                    CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                End If
                CrystalReport1.Formulas(5) = "PATHNO = '" & txt_pathno.Text & "'"
                
                
                
'                If Trim(prepath) = "" Then
'                    CrystalReport1.Formulas(5) = "PATH = '" & "" & " '"
'                Else
'                    CrystalReport1.Formulas(5) = "PATH = '" & " Previous path No. " & Left(prepath, 200) & "'"
'                End If
                CrystalReport1.Formulas(6) = "CAP = '" & "REQUISITION FOR BLOOD PRODUCTS" & " '"
                CrystalReport1.Formulas(7) = "NAME = '" & Trim(txt_name.Text) & " '"
                'CrystalReport1.Action = 1
            End If
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
                Data1.Recordset.MoveNext
            Wend
        End If
End Sub

Public Sub HL()
If PREPATH_FLAG = False Then
    PREPATH_QR(0) = Msk_caseno.FormattedText
    Set PREPATH_RS = PREPATH_QR.OpenResultset(1, 3)
    PREPATH_FLAG = True
Else
    PREPATH_QR(0) = Msk_caseno.FormattedText
    PREPATH_RS.MoreResults
    PREPATH_RS.Requery
End If
If Not PREPATH_RS.EOF Then
    Do While Not PREPATH_RS.EOF
        hemcbc.cmb_preinvest.AddItem req_chglen(PREPATH_RS("REQNNO")) & " " & "$" & PREPATH_RS("PATHNO")
        PREPATH_RS.MoveNext
    Loop
End If
        surgicalPathology.Visible = False
        immunohistochemistry.Visible = False
        frm_hematopathology.Visible = False
        biochemistry.Visible = False
        bceltm.Visible = True
        Label10.Visible = False
        Label11.Visible = False
        cmb_cyprocedure.Visible = False
        cmb_cyprocode.Visible = False
        'cardiology.Visible = False
        Microbiology.Visible = False
        fnac.Visible = False
        CYTOLOGY.Visible = False
        biochemistry.Hide
        Em_lab.Hide
        TUMOURMARKER.Hide
        hemcbc.show
        HEMBF.Hide
        HEMBMA.Hide
        hemcbc.Msk_caseno.Text = Msk_caseno.FormattedText
        hemcbc.txt_name.Text = txt_name.Text
        lbl_req.Left = 3500
        lbl_req.Caption = "REQUISITION - HAEMATOLOGY"
        pathology.Caption = "HematoPathology - HAEMATOLOGY"
End Sub

Public Sub BM()
If PREPATH_FLAG = False Then
    PREPATH_QR(0) = Msk_caseno.FormattedText
    Set PREPATH_RS = PREPATH_QR.OpenResultset(1, 3)
    PREPATH_FLAG = True
Else
    PREPATH_QR(0) = Msk_caseno.FormattedText
    PREPATH_RS.MoreResults
    PREPATH_RS.Requery
End If
If Not PREPATH_RS.EOF Then
    Do While Not PREPATH_RS.EOF
        HEMBMA.cmb_bmaspno.AddItem req_chglen(PREPATH_RS("REQNNO")) & " " & "$" & PREPATH_RS("PATHNO")
        PREPATH_RS.MoveNext
    Loop
End If
        surgicalPathology.Visible = False
        immunohistochemistry.Visible = False
        frm_hematopathology.Visible = False
        biochemistry.Visible = False
        bceltm.Visible = True
        Label10.Visible = False
        Label11.Visible = False
        cmb_cyprocedure.Visible = False
        cmb_cyprocode.Visible = False
        'cardiology.Visible = False
        Microbiology.Visible = False
        fnac.Visible = False
        CYTOLOGY.Visible = FASLE
        biochemistry.Hide
        Em_lab.Hide
        TUMOURMARKER.Hide
        hemcbc.Hide
        HEMBF.Hide
        HEMBMA.show
        HEMBMA.Msk_caseno.Text = Msk_caseno.FormattedText
        HEMBMA.txt_name.Text = txt_name.Text
        lbl_req.Left = 3500
        lbl_req.Caption = "REQUISITION - BONE MARROW ASPIRATION"
        pathology.Caption = "Pathology - BMA"
End Sub

Public Sub BF()
        surgicalPathology.Visible = False
        immunohistochemistry.Visible = False
        frm_hematopathology.Visible = False
        biochemistry.Visible = False
        bceltm.Visible = True
        Label10.Visible = False
        Label11.Visible = False
        cmb_cyprocedure.Visible = False
        cmb_cyprocode.Visible = False
        'cardiology.Visible = False
        Microbiology.Visible = False
        fnac.Visible = False
        CYTOLOGY.Visible = FASLE
        biochemistry.Hide
        Em_lab.Hide
        TUMOURMARKER.Hide
        
        hemcbc.Hide
        HEMBF.show
        HEMBMA.Hide
        HEMBF.Msk_caseno.Text = Msk_caseno.FormattedText
        HEMBF.txt_name.Text = txt_name.Text
        lbl_req.Left = 3500
        lbl_req.Caption = "REQUISITION - BODY FLUIDS"
        pathology.Caption = "Pathology - BODY FLUIDS"
End Sub

Public Sub ED()
        surgicalPathology.Visible = False
        immunohistochemistry.Visible = False
        
        biochemistry.Visible = False
        bceltm.Visible = False
        'cardiology.Visible = False
        Microbiology.Visible = False
        fnac.Visible = False
        CYTOLOGY.Visible = FASLE
        biochemistry.Hide
        Em_lab.Hide
        TUMOURMARKER.Hide
        frm_hematopathology.Visible = True
        frm_hematopathology.Caption = "Endoscopy"
        hemcbc.Hide
        HEMBF.Hide
        HEMBMA.Hide
        cmd_edindication.Visible = False
        'chemistry.Msk_caseno.Text = Msk_caseno.FormattedText
        'biochemistry.txt_name.Text = txt_name.Text
        lbl_req.Left = 3500
        lbl_req.Caption = " ENDOSCOPY - REQUISITION "
        pathology.Caption = "Endoscopy Requisition"
        If txt_modcode.Text = "ED" Then
            lbl_siteed.Visible = True
            txt_edsite.Visible = True
            lbl_biopsycytology.Visible = True
            'cmb_biopsycytology.Visible = True
            lbl_requestremarks.Visible = True
            txt_REQUESTDETAILS.Visible = True
            lbl_edprerequisition.Visible = True
            cmb_edprereqno.Visible = True
            lbl_edprocedureheading.Visible = True
            cmb_edprocedureheading.Visible = True
            list_edts.Visible = True
            EDENABL ("ENA")
            list_edts.clear
            
            list_edts.AddItem "Biopsy"
            list_edts.AddItem "Brushing"
            list_edts.AddItem "FNA"
            list_edts.AddItem "Culture"
            list_edts.AddItem "Aspiration"
            list_edts.AddItem "None"
            
            cmb_edprocedureheading.clear
            cmb_edprocedureheading.AddItem "UPPER G.I. ENDOSCOPY"
            cmb_edprocedureheading.AddItem "SIGMOIDOSCOPY"
            cmb_edprocedureheading.AddItem "COLONOSCOPY"
            cmb_edprocedureheading.AddItem "ERCP"
            cmb_edprocedureheading.AddItem "ENDOSCOPIC ULTRASONOGRAPHY"
            cmb_edprocedureheading.Text = "UPPER G.I. ENDOSCOPY"
    If SELRINVESTFLAG = False Then
        SELRINVEST_QR(0) = Trim(txt_modcode.Text)
        Set SELRINVEST_RS = SELRINVEST_QR.OpenResultset(1, 3)
        SELRINVESTFLAG = True
    Else
        SELRINVEST_QR(0) = Trim(txt_modcode.Text)
        SELRINVEST_RS.MoreResults
        SELRINVEST_RS.Requery
    End If
    If SELRINVEST_RS.RowCount > 0 Then
        cmb_hpprocedure.clear
        cmb_hpprocedurecode.clear
        While Not SELRINVEST_RS.EOF
            cmb_hpprocedurecode.AddItem Trim(SELRINVEST_RS("ser_code"))
            'cbo_proccode.Text = SELRINVEST_RS("ser_code")
            If SELMA101FLAG = False Then
                SELMA101_QR(0) = Trim(SELRINVEST_RS("SER_CODE"))
                Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                SELMA101FLAG = True
             Else
                SELMA101_QR(0) = SELRINVEST_RS("SER_CODE")
                SELMA101_RS.MoreResults
                SELMA101_RS.Requery
            End If
            If SELMA101_RS.RowCount > 0 Then
                While Not SELMA101_RS.EOF
                cmb_hpprocedure.AddItem SELMA101_RS("SER_DET")
                'CMB_CGPROCEDURE.Text = SELMA101_RS("SER_DET")
                SELMA101_RS.MoveNext
                Wend
            End If
            SELRINVEST_RS.MoveNext
        Wend
        
        
    End If
            
        End If
End Sub

Public Sub EC()
        surgicalPathology.Visible = False
        immunohistochemistry.Visible = False
        frm_hematopathology.Visible = True
        frm_hematopathology.Caption = "Cardiology"
        biochemistry.Visible = False
        bceltm.Visible = False
        'cardiology.Visible = False
        Microbiology.Visible = False
        fnac.Visible = False
        CYTOLOGY.Visible = FASLE
        biochemistry.Hide
        Em_lab.Hide
        TUMOURMARKER.Hide
        
        hemcbc.Hide
        HEMBF.Hide
        HEMBMA.Hide
        cmd_edindication.Visible = True
        'chemistry.Msk_caseno.Text = Msk_caseno.FormattedText
        'biochemistry.txt_name.Text = txt_name.Text
        lbl_req.Left = 3500
        lbl_req.Caption = " CARDIOLOGY - REQUISITION "
        pathology.Caption = "Cardiology Requisition"
        ECNEWFLAG = True
        cmb_hpprocedure.Enabled = True
        cmb_hpprocedurecode.Enabled = True
        lbl_hpreqno1.Visible = True
        txt_hpreqno.Visible = True
        If Trim(txt_modcode.Text) = "2D" Or Trim(txt_modcode.Text) = "PF" Or Trim(txt_modcode.Text) = "ST" Then
            lbl_siteed.Visible = False
            txt_edsite.Visible = False
            lbl_biopsycytology.Visible = False
            'cmb_biopsycytology.Visible = False
            lbl_requestremarks.Visible = False
            txt_REQUESTDETAILS.Visible = False
            lbl_edprerequisition.Visible = False
            cmb_edprereqno.Visible = False
            lbl_edprocedureheading.Visible = False
            cmb_edprocedureheading.Visible = False
            list_edts.Visible = False
            
            cmd_edindication.Enabled = True
            
        Else
            cmd_edindication.Enabled = False
            lbl_siteed.Visible = False
            txt_edsite.Visible = False
            lbl_biopsycytology.Visible = False
'            cmb_biopsycytology.Visible = False
            lbl_requestremarks.Visible = False
            txt_REQUESTDETAILS.Visible = False
            lbl_edprerequisition.Visible = False
            cmb_edprereqno.Visible = False
            lbl_edprocedureheading.Visible = False
            cmb_edprocedureheading.Visible = False
            list_edts.Visible = False
            
        End If
        
    If SELRINVESTFLAG = False Then
        SELRINVEST_QR(0) = Trim(txt_modcode.Text)
        Set SELRINVEST_RS = SELRINVEST_QR.OpenResultset(1, 3)
        SELRINVESTFLAG = True
    Else
        SELRINVEST_QR(0) = Trim(txt_modcode.Text)
        SELRINVEST_RS.MoreResults
        SELRINVEST_RS.Requery
    End If
    If SELRINVEST_RS.RowCount > 0 Then
        cmb_hpprocedure.clear
        cmb_hpprocedurecode.clear
        If Trim(txt_modcode.Text) = "EK" Then
            While Not SELRINVEST_RS.EOF
                cmb_hpprocedurecode.AddItem Trim(SELRINVEST_RS("ser_code"))
                cmb_hpprocedurecode.Text = SELRINVEST_RS("ser_code")
                If SELMA101FLAG = False Then
                    SELMA101_QR(0) = Trim(SELRINVEST_RS("SER_CODE"))
                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                    SELMA101FLAG = True
                 Else
                    SELMA101_QR(0) = SELRINVEST_RS("SER_CODE")
                    SELMA101_RS.MoreResults
                    SELMA101_RS.Requery
                End If
                If SELMA101_RS.RowCount > 0 Then
                    While Not SELMA101_RS.EOF
                    cmb_hpprocedure.AddItem SELMA101_RS("SER_DET")
                    cmb_hpprocedure.Text = SELMA101_RS("SER_DET")
                    SELMA101_RS.MoveNext
                    Wend
                End If
                SELRINVEST_RS.MoveNext
            Wend
        End If
        If Trim(txt_modcode.Text) = "PF" Then
            While Not SELRINVEST_RS.EOF
                cmb_hpprocedurecode.AddItem Trim(SELRINVEST_RS("ser_code"))
                cmb_hpprocedurecode.Text = SELRINVEST_RS("ser_code")
                If SELMA101FLAG = False Then
                    SELMA101_QR(0) = Trim(SELRINVEST_RS("SER_CODE"))
                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                    SELMA101FLAG = True
                 Else
                    SELMA101_QR(0) = SELRINVEST_RS("SER_CODE")
                    SELMA101_RS.MoreResults
                    SELMA101_RS.Requery
                End If
                If SELMA101_RS.RowCount > 0 Then
                    While Not SELMA101_RS.EOF
                    cmb_hpprocedure.AddItem SELMA101_RS("SER_DET")
                    cmb_hpprocedure.Text = SELMA101_RS("SER_DET")
                    SELMA101_RS.MoveNext
                    Wend
                End If
                SELRINVEST_RS.MoveNext
            Wend
            cmb_hpprocedurecode.Text = "K107"
            cmb_hpprocedure.Text = "PFT"
        End If
        If Trim(txt_modcode.Text) = "ST" Then
            While Not SELRINVEST_RS.EOF
                cmb_hpprocedurecode.AddItem Trim(SELRINVEST_RS("ser_code"))
                
                If SELMA101FLAG = False Then
                    SELMA101_QR(0) = Trim(SELRINVEST_RS("SER_CODE"))
                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                    SELMA101FLAG = True
                 Else
                    SELMA101_QR(0) = SELRINVEST_RS("SER_CODE")
                    SELMA101_RS.MoreResults
                    SELMA101_RS.Requery
                End If
                If SELMA101_RS.RowCount > 0 Then
                    While Not SELMA101_RS.EOF
                    cmb_hpprocedure.AddItem SELMA101_RS("SER_DET")
                    
                    SELMA101_RS.MoveNext
                    Wend
                End If
                SELRINVEST_RS.MoveNext
            Wend
            cmb_hpprocedurecode.Text = "K109"
            cmb_hpprocedure.Text = "STRESS TEST"
        End If
        If Trim(txt_modcode.Text) = "2D" Then
            While Not SELRINVEST_RS.EOF
                cmb_hpprocedurecode.AddItem Trim(SELRINVEST_RS("ser_code"))
                
                If SELMA101FLAG = False Then
                    SELMA101_QR(0) = Trim(SELRINVEST_RS("SER_CODE"))
                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                    SELMA101FLAG = True
                 Else
                    SELMA101_QR(0) = SELRINVEST_RS("SER_CODE")
                    SELMA101_RS.MoreResults
                    SELMA101_RS.Requery
                End If
                If SELMA101_RS.RowCount > 0 Then
                    While Not SELMA101_RS.EOF
                    cmb_hpprocedure.AddItem SELMA101_RS("SER_DET")
                    
                    SELMA101_RS.MoveNext
                    Wend
                End If
                SELRINVEST_RS.MoveNext
            Wend
            'Next 2 statements are commented by Shraddha Pandit on 04/02/09
            'cmb_hpprocedurecode.Text = "K102"
            'cmb_hpprocedure.Text = "ECHOCARDIOGRAM"
        End If
    End If

End Sub

Public Sub EDCLEAR()
cmb_hpprocedure.clear
cmb_hpprocedurecode.clear
cmb_edprereqno.clear
cmb_edprocedureheading.clear
txt_edsite.Text = ""
txt_REQUESTDETAILS.Text = ""
list_edts.clear
End Sub

Public Sub EDENABL(ST As String)
If ST = "ENA" Then
    cmb_hpprocedure.Enabled = True
    cmb_hpprocedurecode.Enabled = True
    cmb_edprereqno.Enabled = True
    cmb_edprocedureheading.Enabled = True
    txt_edsite.Enabled = True
    txt_REQUESTDETAILS.Enabled = True
    list_edts.Enabled = True
    txt_hpreqno.Enabled = False
End If
If ST = "DIS" Then
    cmb_hpprocedure.Enabled = False
    cmb_hpprocedurecode.Enabled = False
    cmb_edprereqno.Enabled = False
    cmb_edprocedureheading.Enabled = False
    txt_edsite.Enabled = False
    txt_REQUESTDETAILS.Enabled = False
    list_edts.Enabled = False
    txt_hpreqno.Enabled = False
End If
If ST = "MOD" Then
    cmb_hpprocedure.Enabled = False
    cmb_hpprocedurecode.Enabled = False
    cmb_edprereqno.Enabled = False
    cmb_edprocedureheading.Enabled = True
    txt_edsite.Enabled = True
    txt_REQUESTDETAILS.Enabled = True
    list_edts.Enabled = True
    txt_hpreqno.Enabled = False
End If
End Sub

Public Sub EDASSIGN()
Dim durill As String
Dim I As Integer

On Error GoTo MSG:
'*** INSERTING INT0 RREQUISITION
INSRREQUISITN_QR(0) = req_chglen(Trim(txt_hpreqno.Text))
INSRREQUISITN_QR(1) = GET_REVDATE(CStr(DATE11))
INSRREQUISITN_QR(2) = Trim(txt_time.Text)
INSRREQUISITN_QR(3) = Trim(Msk_caseno.FormattedText)
INSRREQUISITN_QR(4) = Trim(txt_dept.Text)
INSRREQUISITN_QR(5) = Trim(txt_modcode.Text)
INSRREQUISITN_QR(6) = Trim(txt_status.Text)
INSRREQUISITN_QR(7) = Trim(txt_catgory.Text)
INSRREQUISITN_QR(8) = Trim(cmb_hpprocedurecode.Text)
INSRREQUISITN_QR(9) = UCase(Trim(cmb_refdoc.Text))
'INSRREQUISITN_QR(9) = Trim(txt_pridocser.Text)

'Added by Milind on 27-03-2006 for inserting date for the samples collected at outside
    If Check1.Value = 1 Then
        Dim Dt As String
        Dt = "18/02/1980"
        INSRREQUISITN_QR(15) = Format(CDate(Dt), "dd/mm/yyyy")
    Else
        INSRREQUISITN_QR(15) = Null
    End If
'upto here

    If SELRREQUISITNFLAG = False Then
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
        SELRREQUISITNFLAG = True
    Else
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        SELRREQUISITN_RS.MoreResults
        SELRREQUISITN_RS.Requery
    End If
    If SELRREQUISITN_RS.RowCount > 0 Then
        INSRREQUISITN_QR(10) = "YES"
    Else
        INSRREQUISITN_QR(10) = "NO"
    End If
If opt_patient.Value = True Then
    NO_CHARGE
    ''' Added by Sandhya dated 26.6.2009
    ''' for NC category & charged patient
    If txt_catgory.Text = "NC" And No_Chargeflag = True Then
        INSRREQUISITN_QR(11) = "NOCHARGE"
        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
        INSRREQUISITN_QR(13) = "P"
    'No_Chargeflag = False
    ElseIf txt_catgory.Text = "C" Then
        If cmb_hpprocedurecode.Text = "M011" Or cmb_hpprocedurecode.Text = "M012" Or cmb_hpprocedurecode.Text = "M013" Or cmb_hpprocedurecode.Text = "M014" Or cmb_hpprocedurecode.Text = "M015" Or cmb_hpprocedurecode.Text = "M701" Or cmb_hpprocedurecode.Text = "M702" Then
            INSRREQUISITN_QR(11) = Null
            INSRREQUISITN_QR(12) = Null
            INSRREQUISITN_QR(13) = Null
        Else
            INSRREQUISITN_QR(11) = "NOCHARGE"
            INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
            INSRREQUISITN_QR(13) = "P"
        End If
    Else
        INSRREQUISITN_QR(11) = Null
        INSRREQUISITN_QR(12) = Null
        INSRREQUISITN_QR(13) = Null
    End If
    
End If
If opt_staff.Value = True Then
        INSRREQUISITN_QR(11) = "NOCHARGE"
        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
        INSRREQUISITN_QR(13) = "P"
End If
INSRREQUISITN_QR(14) = 0

INSRREQUISITN_QR.Execute
    '*** INSERTING INT0 RREQUISITION
    '*** TISSUE SAMPLING
    TSSAMPLE = ""
    If list_edts.ListCount <> 0 Then
        For I = 0 To list_edts.ListCount - 1
            
            If list_edts.Selected(I) = True Then
               TSSAMPLE = TSSAMPLE & list_edts.List(I) & " + "
            End If
        Next I
        TSSAMPLE = Left(TSSAMPLE, Len(Trim(TSSAMPLE)) - 1)
    End If
    '*** TISSUE SAMPLING
    
    '**** INSERTING INTO DREQUISITION

INSDREQUISITN_QR(0) = req_chglen(Trim(txt_hpreqno.Text))
INSDREQUISITN_QR(1) = ""
INSDREQUISITN_QR(2) = Trim(Msk_caseno.FormattedText)
INSDREQUISITN_QR(3) = Null
INSDREQUISITN_QR(4) = Null
dur_ill = txt_year & "/" & Trim(CMB_DURILLMONTH.Text)
INSDREQUISITN_QR(5) = dur_ill
INSDREQUISITN_QR(6) = cmb_edprocedureheading.Text
INSDREQUISITN_QR(7) = Null
INSDREQUISITN_QR(8) = Trim(TSSAMPLE)
INSDREQUISITN_QR(9) = Null
INSDREQUISITN_QR(10) = Null
INSDREQUISITN_QR(11) = Null
INSDREQUISITN_QR(12) = Null
INSDREQUISITN_QR(13) = Null
INSDREQUISITN_QR(14) = Trim(cmb_refdoc.Text)
INSDREQUISITN_QR(15) = Trim(cmb_reftodoc.Text)
INSDREQUISITN_QR(16) = Null
If Trim(txt_status.Text) = "IP" Then
    INSDREQUISITN_QR(17) = Trim(txt_ward.Text)
    INSDREQUISITN_QR(18) = Trim(txt_bedno.Text)
Else
    INSDREQUISITN_QR(17) = Null
    INSDREQUISITN_QR(18) = Null
End If
INSDREQUISITN_QR(19) = Trim(txt_REQUESTDETAILS.Text)
If Len(Trim(txt_edsite.Text)) > 0 Then
    INSDREQUISITN_QR(20) = Null
    INSDREQUISITN_QR(20) = Trim(txt_edsite.Text)
Else
    INSDREQUISITN_QR(20) = Null
End If
INSDREQUISITN_QR(21) = Null
INSDREQUISITN_QR.Execute
'**** INSERTING INTO DREQUISITION

 '*** UPDATING THE PROVISIONAL DIAGNOSIS
 
If opt_patient.Value = True Then
If SELTA1012FLAG = False Then
    selta1012_qr(0) = Msk_caseno.FormattedText
    Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
    SELTA1012FLAG = True
Else
     selta1012_qr(0) = Msk_caseno.FormattedText
     selta1012_rs.MoreResults
     selta1012_rs.Requery
End If
If selta1012_rs.RowCount > 0 Then
    UPDTA1012_QR(0) = Trim(txt_indctinvstg.Text)
    UPDTA1012_QR(1) = Msk_caseno.FormattedText
    UPDTA1012_QR.Execute
Else
    insta1012_qr(0) = Msk_caseno.FormattedText
    insta1012_qr(1) = Trim(txt_indctinvstg.Text)
    insta1012_qr.Execute
End If
End If
'*** UPDATING THE PROVISIONAL DIAGNOSIS
'*** INSERTING INTO DMULPROCE FOR COLLECTION AND DEPT COLLECTION SCREEN
''    INSDMULPROCE_QR(0) = req_chglen(txt_spreqno.Text)
''    INSDMULPROCE_QR(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
''    INSDMULPROCE_QR(2) = Msk_caseno.FormattedText
''    INSDMULPROCE_QR(3) = cmb_procedurecode.Text
''    INSDMULPROCE_QR.Execute

'***INSERTING INTO DMULPROCE FOR COLLECTION AND DEPT COLLECTION SCREEN
'***INSERTING INTO DREQSTAFFD TABLE
If opt_staff.Value = True Then
    INSDREQSTAFFD_QR(0) = req_chglen(txt_hpreqno.Text)
    INSDREQSTAFFD_QR(1) = Msk_caseno.Text
    INSDREQSTAFFD_QR(2) = Trim(txt_name.Text)
    INSDREQSTAFFD_QR(3) = Trim(txt_Age)
    INSDREQSTAFFD_QR(4) = Trim(txt_sex)
    If Len(Trim(cmb_dependents.Text)) = 0 Then
        INSDREQSTAFFD_QR(5) = "SELF"
    Else
        I = InStr(cmb_dependents, "$")
        INSDREQSTAFFD_QR(5) = Mid(cmb_dependents, I + 1, 10)
    End If
    INSDREQSTAFFD_QR.Execute
End If
'***INSERTING INTO DREQSTAFFD TABLE
    If SELCLINHISTFLAG = False Then
        SELCLINHIST_QR(0) = Msk_caseno.FormattedText
        Set SELCLINHIST_RS = SELCLINHIST_QR.OpenResultset(1, 3)
        SELCLINHISTFLAG = True
    Else
        SELCLINHIST_QR(0) = Msk_caseno.FormattedText
        SELCLINHIST_RS.MoreResults
        SELCLINHIST_RS.Requery
    End If
'    If SELCLINHIST_RS.RowCount > 0 Then
'        updclinhist_qr(0) = Trim(txt_clinhist.Text)
'        updclinhist_qr(1) = Trim(Msk_caseno.Text)
'        updclinhist_qr.Execute
'    Else
        insclinhist_qr(0) = Trim(Msk_caseno.FormattedText)
        insclinhist_qr(1) = GET_REVDATE(Trim(DATE11))
        insclinhist_qr(2) = Trim(time1)
        If Trim(txt_clinhist.Text) = "NYD" Then
            insclinhist_qr(3) = "NYD"
        Else
            insclinhist_qr(3) = Trim(txt_clinhist.Text)
        End If
        insclinhist_qr.Execute
    'End If

Data1.UpdateRecord
Data1.Recordset.Bookmark = Data1.Recordset.LastModified
Exit Sub
MSG:
MsgBox Err.Description

End Sub

Public Sub ECASSIGN()
Dim durill As String
Dim I As Integer

On Error GoTo MSG:
'*** INSERTING INT0 RREQUISITION
INSRREQUISITN_QR(0) = req_chglen(Trim(txt_hpreqno.Text))
INSRREQUISITN_QR(1) = GET_REVDATE(CStr(DATE11))
INSRREQUISITN_QR(2) = Trim(txt_time.Text)
INSRREQUISITN_QR(3) = Trim(Msk_caseno.FormattedText)
INSRREQUISITN_QR(4) = Trim(txt_dept.Text)
INSRREQUISITN_QR(5) = Trim(txt_modcode.Text)
INSRREQUISITN_QR(6) = Trim(txt_status.Text)
INSRREQUISITN_QR(7) = Trim(txt_catgory.Text)
INSRREQUISITN_QR(8) = Trim(cmb_hpprocedurecode.Text)
INSRREQUISITN_QR(9) = UCase(Trim(cmb_refdoc.Text))
'INSRREQUISITN_QR(9) = Trim(txt_pridocser.Text)

'Added by Milind on 27-03-2006 for inserting date for the samples collected at outside
    If Check1.Value = 1 Then
        Dim Dt As String
        Dt = "18/02/1980"
        INSRREQUISITN_QR(15) = Format(CDate(Dt), "dd/mm/yyyy")
    Else
        INSRREQUISITN_QR(15) = Null
    End If
'upto here

    If SELRREQUISITNFLAG = False Then
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
        SELRREQUISITNFLAG = True
    Else
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        SELRREQUISITN_RS.MoreResults
        SELRREQUISITN_RS.Requery
    End If
    If SELRREQUISITN_RS.RowCount > 0 Then
        INSRREQUISITN_QR(10) = "YES"
    Else
        INSRREQUISITN_QR(10) = "NO"
    End If
If opt_patient.Value = True Then
    NO_CHARGE
    ''' Added by Sandhya dated 26.6.2009
    ''' for NC category & charged patient
    If txt_catgory.Text = "NC" And No_Chargeflag = True Then
        INSRREQUISITN_QR(11) = "NOCHARGE"
        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
        INSRREQUISITN_QR(13) = "E"
        'No_Chargeflag = False
    ''ICU
    ElseIf lbl_icureq.Visible = True Then
        'If Trim(CMB_ICU1.Text) = "YES" Then
            INSRREQUISITN_QR(11) = Null
            INSRREQUISITN_QR(12) = Null
            'INSRREQUISITN_QR(13) = "E"
            'INSRREQUISITN_QR(13) = " "
             INSRREQUISITN_QR(13) = Null
        'End If
    ''ICU
    Else
        
        INSRREQUISITN_QR(11) = Null
        INSRREQUISITN_QR(12) = Null
        INSRREQUISITN_QR(13) = Null
       
    End If
End If
If opt_staff.Value = True Then
        INSRREQUISITN_QR(11) = "NOCHARGE"
        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
        INSRREQUISITN_QR(13) = "E"
End If
INSRREQUISITN_QR(14) = 0

INSRREQUISITN_QR.Execute
    '*** INSERTING INT0 RREQUISITION
    
    
    '**** INSERTING INTO DREQUISITION

INSDREQUISITN_QR(0) = req_chglen(Trim(txt_hpreqno.Text))
INSDREQUISITN_QR(1) = ""
INSDREQUISITN_QR(2) = Trim(Msk_caseno.FormattedText)
INSDREQUISITN_QR(3) = Null
INSDREQUISITN_QR(4) = Null
dur_ill = txt_year & "/" & Trim(CMB_DURILLMONTH.Text)
INSDREQUISITN_QR(5) = dur_ill
INSDREQUISITN_QR(6) = Null
INSDREQUISITN_QR(7) = Null
INSDREQUISITN_QR(8) = Null
INSDREQUISITN_QR(9) = Null
INSDREQUISITN_QR(10) = Null
INSDREQUISITN_QR(11) = Null
INSDREQUISITN_QR(12) = Null
INSDREQUISITN_QR(13) = Null
INSDREQUISITN_QR(14) = Trim(cmb_refdoc.Text)
INSDREQUISITN_QR(15) = Trim(cmb_reftodoc.Text)
INSDREQUISITN_QR(16) = Null
If Trim(txt_status.Text) = "IP" Then
    INSDREQUISITN_QR(17) = Trim(txt_ward.Text)
    INSDREQUISITN_QR(18) = Trim(txt_bedno.Text)
Else
    INSDREQUISITN_QR(17) = Null
    INSDREQUISITN_QR(18) = Null
End If
INSDREQUISITN_QR(19) = Null
INSDREQUISITN_QR(20) = Null
INSDREQUISITN_QR(21) = Null
INSDREQUISITN_QR.Execute
'**** INSERTING INTO DREQUISITION
'**** INSERTING INTO INDICATION
INSINDICATIONQR(0) = req_chglen(txt_hpreqno.Text)
If Len(Trim(ind1)) > 0 Then
    INSINDICATIONQR(1) = Null
    INSINDICATIONQR(1) = ind1
Else
    INSINDICATIONQR(1) = Null
End If
If Len(Trim(ind2)) > 0 Then
    INSINDICATIONQR(2) = Null
    INSINDICATIONQR(2) = ind2
Else
    INSINDICATIONQR(2) = Null
End If
If Len(Trim(ind3)) > 0 Then
    INSINDICATIONQR(3) = Null
    INSINDICATIONQR(3) = ind3
Else
    INSINDICATIONQR(3) = Null
End If
If Len(Trim(ind4)) > 0 Then
    INSINDICATIONQR(4) = Null
    INSINDICATIONQR(4) = ind4
Else
    INSINDICATIONQR(4) = Null
End If
If Len(Trim(ind5)) > 0 Then
    INSINDICATIONQR(5) = Null
    INSINDICATIONQR(5) = ind5
Else
    INSINDICATIONQR(5) = Null
End If
If Len(Trim(ind6)) > 0 Then
    INSINDICATIONQR(6) = Null
    INSINDICATIONQR(6) = ind6
Else
    INSINDICATIONQR(6) = ind6
End If
If Len(Trim(ind7)) > 0 Then
    INSINDICATIONQR(7) = Null
    INSINDICATIONQR(7) = ind7
Else
    INSINDICATIONQR(7) = Null
End If
If Len(Trim(ind8)) > 0 Then
    INSINDICATIONQR(8) = Null
    INSINDICATIONQR(8) = ind8
Else
    INSINDICATIONQR(8) = Null
End If
If Len(Trim(ind9)) > 0 Then
    INSINDICATIONQR(9) = Null
    INSINDICATIONQR(9) = ind9
Else
    INSINDICATIONQR(9) = Null
End If
If Len(Trim(tp1)) > 0 Then
    INSINDICATIONQR(10) = Null
    INSINDICATIONQR(10) = tp1
Else
    INSINDICATIONQR(10) = Null
End If
If Len(Trim(tp2)) > 0 Then
    INSINDICATIONQR(11) = Null
    INSINDICATIONQR(11) = tp2
Else
    INSINDICATIONQR(11) = Null
End If

If Len(Trim(tp3)) > 0 Then
    INSINDICATIONQR(12) = Null
    INSINDICATIONQR(12) = tp3
Else
    INSINDICATIONQR(12) = Null
End If

INSINDICATIONQR(13) = drug
INSINDICATIONQR(14) = dose
INSINDICATIONQR(15) = noofcycles
INSINDICATIONQR(16) = bsa
If Len(Trim(OTHERS)) > 0 Then
    INSINDICATIONQR(17) = Null
    INSINDICATIONQR(17) = OTHERS
Else
    INSINDICATIONQR(17) = Null
End If
INSINDICATIONQR.Execute
'**** INSERTING INTO INDICATION
 '*** UPDATING THE PROVISIONAL DIAGNOSIS
 
If opt_patient.Value = True Then
If SELTA1012FLAG = False Then
    selta1012_qr(0) = Msk_caseno.FormattedText
    Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
    SELTA1012FLAG = True
Else
     selta1012_qr(0) = Msk_caseno.FormattedText
     selta1012_rs.MoreResults
     selta1012_rs.Requery
End If
If selta1012_rs.RowCount > 0 Then
    UPDTA1012_QR(0) = Trim(txt_indctinvstg.Text)
    UPDTA1012_QR(1) = Msk_caseno.FormattedText
    UPDTA1012_QR.Execute
Else
    insta1012_qr(0) = Msk_caseno.FormattedText
    insta1012_qr(1) = Trim(txt_indctinvstg.Text)
    insta1012_qr.Execute
End If
End If
'*** UPDATING THE PROVISIONAL DIAGNOSIS
'*** INSERTING INTO DMULPROCE FOR COLLECTION AND DEPT COLLECTION SCREEN
''    INSDMULPROCE_QR(0) = req_chglen(txt_spreqno.Text)
''    INSDMULPROCE_QR(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
''    INSDMULPROCE_QR(2) = Msk_caseno.FormattedText
''    INSDMULPROCE_QR(3) = cmb_procedurecode.Text
''    INSDMULPROCE_QR.Execute

'***INSERTING INTO DMULPROCE FOR COLLECTION AND DEPT COLLECTION SCREEN
'***INSERTING INTO DREQSTAFFD TABLE
If opt_staff.Value = True Then
    INSDREQSTAFFD_QR(0) = req_chglen(txt_hpreqno.Text)
    INSDREQSTAFFD_QR(1) = Msk_caseno.Text
    INSDREQSTAFFD_QR(2) = Trim(txt_name.Text)
    INSDREQSTAFFD_QR(3) = Trim(txt_Age)
    INSDREQSTAFFD_QR(4) = Trim(txt_sex)
    If Len(Trim(cmb_dependents.Text)) = 0 Then
        INSDREQSTAFFD_QR(5) = "SELF"
    Else
        I = InStr(cmb_dependents, "$")
        INSDREQSTAFFD_QR(5) = Mid(cmb_dependents, I + 1, 10)
    End If
    INSDREQSTAFFD_QR.Execute
End If
'***INSERTING INTO DREQSTAFFD TABLE
    If SELCLINHISTFLAG = False Then
        SELCLINHIST_QR(0) = Msk_caseno.FormattedText
        Set SELCLINHIST_RS = SELCLINHIST_QR.OpenResultset(1, 3)
        SELCLINHISTFLAG = True
    Else
        SELCLINHIST_QR(0) = Msk_caseno.FormattedText
        SELCLINHIST_RS.MoreResults
        SELCLINHIST_RS.Requery
    End If
'    If SELCLINHIST_RS.RowCount > 0 Then
'        updclinhist_qr(0) = Trim(txt_clinhist.Text)
'        updclinhist_qr(1) = Trim(Msk_caseno.Text)
'        updclinhist_qr.Execute
'    Else
        insclinhist_qr(0) = Trim(Msk_caseno.FormattedText)
        insclinhist_qr(1) = GET_REVDATE(Trim(DATE11))
        insclinhist_qr(2) = Trim(time1)
        If Trim(txt_clinhist.Text) = "NYD" Then
            insclinhist_qr(3) = "NYD"
        Else
            insclinhist_qr(3) = Trim(txt_clinhist.Text)
        End If
        insclinhist_qr.Execute
    'End If
''UPDATE DREQUISITN TABLE
If SELDISICU_FLAG = False Then
    SELDISICU_QR(0) = Trim(ipadd)
    Set SELDISICU_RS = SELDISICU_QR.OpenResultset(1, 3)
    SELDISICU_FLAG = True
Else
    SELDISICU_QR(0) = Trim(ipadd)
    SELDISICU_RS.MoreResults
    SELDISICU_RS.Requery
End If
If Not SELDISICU_RS.EOF Then
    UPDICU_QR(0) = "YES"
    UPDICU_QR(1) = req_chglen(Trim(txt_hpreqno.Text))
    UPDICU_QR.Execute
End If
    
''UPDATE DREQUISITN TABLE
Data1.UpdateRecord
Data1.Recordset.Bookmark = Data1.Recordset.LastModified
Exit Sub
MSG:
MsgBox Err.Description

End Sub

Public Sub ECVIEW()
    If Trim(txt_modcode.Text) = "2D" Then
        cardiologyind.show
    End If
        'EDVIEWFLAG = True
    If Trim(txt_modcode.Text) = "2D" Then
        cardiologyind.echoi.Visible = True
        cardiologyind.pet.Visible = False
        cardiologyind.sti.Visible = False
        If Trim(SELINDICATIONRS("IND1")) = cardiologyind.CHK_2DIND1.Caption Then
            cardiologyind.CHK_2DIND1.Value = 1
        Else
            cardiologyind.CHK_2DIND1.Value = 0
        End If
        If Trim(SELINDICATIONRS("IND2")) = cardiologyind.CHK_2DIND2.Caption Then
            cardiologyind.CHK_2DIND2.Value = 1
        Else
            cardiologyind.CHK_2DIND2.Value = 0
        End If
        If Trim(SELINDICATIONRS("IND3")) = cardiologyind.CHK_2DIND3.Caption Then
            cardiologyind.CHK_2DIND3.Value = 1
        Else
            cardiologyind.CHK_2DIND3.Value = 0
        End If
        If Trim(SELINDICATIONRS("IND4")) = cardiologyind.CHK_2DIND4.Caption Then
            cardiologyind.CHK_2DIND4.Value = 1
        Else
            cardiologyind.CHK_2DIND4.Value = 0
        End If
        If Trim(SELINDICATIONRS("IND5")) = cardiologyind.CHK_2DIND5.Caption Then
            cardiologyind.CHK_2DIND5.Value = 1
        Else
            cardiologyind.CHK_2DIND5.Value = 0
        End If
        If Trim(SELINDICATIONRS("IND6")) = cardiologyind.chk_2DIND6.Caption Then
            cardiologyind.chk_2DIND6.Value = 1
        Else
            cardiologyind.chk_2DIND6.Value = 0
        End If
        If Trim(SELINDICATIONRS("IND7")) = cardiologyind.chk_2DIND7.Caption Then
            cardiologyind.chk_2DIND7.Value = 1
        Else
            cardiologyind.chk_2DIND7.Value = 0
        End If
        If Trim(SELINDICATIONRS("IND8")) = cardiologyind.CHK_2DIND8.Caption Then
            cardiologyind.CHK_2DIND8.Value = 1
        Else
            cardiologyind.CHK_2DIND8.Value = 0
        End If
        If Trim(SELINDICATIONRS("TP1")) = cardiologyind.chk_Surgery.Caption Then
            cardiologyind.chk_Surgery.Value = 1
        Else
            cardiologyind.chk_Surgery.Value = 0
        End If
        If Trim(SELINDICATIONRS("TP2")) = cardiologyind.chk_chemotherapy.Caption Then
            cardiologyind.chk_chemotherapy.Value = 1
        Else
            cardiologyind.chk_chemotherapy.Value = 0
        End If
        
        If Trim(SELINDICATIONRS("TP3")) = cardiologyind.chk_radiotherapy.Caption Then
            cardiologyind.chk_radiotherapy.Value = 1
        Else
            cardiologyind.chk_radiotherapy.Value = 0
        End If
        cardiologyind.txt_drug.Text = SELINDICATIONRS("DRUG")
        cardiologyind.txt_dose.Text = SELINDICATIONRS("DOSE")
        cardiologyind.txt_noofcycles.Text = SELINDICATIONRS("NOOFCYCLES")
        cardiologyind.txt_bsa.Text = SELINDICATIONRS("BSA")
        'cardiologyind.txt_lastcycle = SELINDICATIONRS("")
        If SELINDICATIONRS.EOF = False Then
            cardiologyind.TXT_OTHERS = " " & SELINDICATIONRS("OTHERS")
        End If
    End If
    If Trim(txt_modcode.Text) = "PF" Then
        cardiologyind.echoi.Visible = False
        cardiologyind.pet.Visible = True
        cardiologyind.sti.Visible = False
        If Trim(SELINDICATIONRS("IND1")) = cardiologyind.CHK_PFIND1.Caption Then
            cardiologyind.CHK_PFIND1.Value = 1
        Else
            cardiologyind.CHK_PFIND1.Value = 0
        End If
        If Trim(SELINDICATIONRS("IND2")) = cardiologyind.CHK_PFIND2.Caption Then
            cardiologyind.CHK_PFIND2.Value = 1
        Else
            cardiologyind.CHK_PFIND2.Value = 0
        End If
        If Trim(SELINDICATIONRS("IND3")) = cardiologyind.CHK_PFIND3.Caption Then
            cardiologyind.CHK_PFIND3.Value = 1
        Else
            cardiologyind.CHK_PFIND3.Value = 0
        End If
        If Trim(SELINDICATIONRS("IND4")) = cardiologyind.CHK_PFIND4.Caption Then
            cardiologyind.CHK_2DIND4.Value = 1
        Else
            cardiologyind.CHK_PFIND4.Value = 0
        End If
        If Trim(SELINDICATIONRS("IND5")) = cardiologyind.CHK_PFIND5.Caption Then
            cardiologyind.CHK_PFIND5.Value = 1
        Else
            cardiologyind.CHK_PFIND5.Value = 0
        End If
        If Trim(SELINDICATIONRS("IND6")) = cardiologyind.CHK_PFIND6.Caption Then
            cardiologyind.CHK_PFIND6.Value = 1
        Else
            cardiologyind.CHK_PFIND6.Value = 0
        End If
        If Trim(SELINDICATIONRS("IND7")) = cardiologyind.CHK_PFIND7.Caption Then
            cardiologyind.CHK_PFIND7.Value = 1
        Else
            cardiologyind.CHK_PFIND7.Value = 0
        End If
'        If Trim(SELINDICATIONRS("IND8")) = cardiologyind.CHK_PF.Caption Then
'            cardiologyind.CHK_PFIND8.Value = 1
'        Else
'            cardiologyind.CHK_PFIND8.Value = 0
'        End If
        If Trim(SELINDICATIONRS("TP1")) = cardiologyind.chk_Surgery.Caption Then
            cardiologyind.chk_Surgery.Value = 1
        Else
            cardiologyind.chk_Surgery.Value = 0
        End If
        If Trim(SELINDICATIONRS("TP2")) = cardiologyind.chk_chemotherapy.Caption Then
            cardiologyind.chk_chemotherapy.Value = 1
        Else
            cardiologyind.chk_chemotherapy.Value = 0
        End If
        
        If Trim(SELINDICATIONRS("TP3")) = cardiologyind.chk_radiotherapy.Caption Then
            cardiologyind.chk_radiotherapy.Value = 1
        Else
            cardiologyind.chk_radiotherapy.Value = 0
        End If
        
       
    End If
    If Trim(txt_modcode.Text) = "ST" Then
        
        cardiologyind.echoi.Visible = False
        cardiologyind.pet.Visible = False
        cardiologyind.sti.Visible = True
        If Trim(SELINDICATIONRS("IND1")) = cardiologyind.CHK_STIND1.Caption Then
            cardiologyind.CHK_STIND1.Value = 1
        Else
            cardiologyind.CHK_STIND1.Value = 0
        End If
        If Trim(SELINDICATIONRS("IND2")) = cardiologyind.CHK_STIND2.Caption Then
            cardiologyind.CHK_STIND2.Value = 1
        Else
            cardiologyind.CHK_STIND2.Value = 0
        End If
        If Trim(SELINDICATIONRS("IND3")) = cardiologyind.CHK_STIND3.Caption Then
            cardiologyind.CHK_STIND3.Value = 1
        Else
            cardiologyind.CHK_STIND3.Value = 0
        End If
        If Trim(SELINDICATIONRS("IND4")) = cardiologyind.CHK_STIND4.Caption Then
            cardiologyind.CHK_STIND4.Value = 1
        Else
            cardiologyind.CHK_STIND4.Value = 0
        End If
        If Trim(SELINDICATIONRS("IND5")) = cardiologyind.CHK_STIND5.Caption Then
            cardiologyind.CHK_STIND5.Value = 1
        Else
            cardiologyind.CHK_STIND5.Value = 0
        End If
        If Trim(SELINDICATIONRS("IND6")) = cardiologyind.CHK_STIND6.Caption Then
            cardiologyind.CHK_STIND6.Value = 1
        Else
            cardiologyind.CHK_STIND6.Value = 0
        End If
        
        
        If Trim(SELINDICATIONRS("TP1")) = cardiologyind.chk_Surgery.Caption Then
            cardiologyind.chk_Surgery.Value = 1
        Else
            cardiologyind.chk_Surgery.Value = 0
        End If
        If Trim(SELINDICATIONRS("TP2")) = cardiologyind.chk_chemotherapy.Caption Then
            cardiologyind.chk_chemotherapy.Value = 1
        Else
            cardiologyind.chk_chemotherapy.Value = 0
        End If
        
        If Trim(SELINDICATIONRS("TP3")) = cardiologyind.chk_radiotherapy.Caption Then
            cardiologyind.chk_radiotherapy.Value = 1
        Else
            cardiologyind.chk_radiotherapy.Value = 0
        End If
        If SELINDICATIONRS.EOF = False Then
            cardiologyind.txt_stothers = " " & SELINDICATIONRS("OTHERS")
        End If
    End If
End Sub



Public Sub EDUPD()
'**** Updating the drequisition table
        upddrequisitn_qr(0) = (cmb_refdoc.Text)
        upddrequisitn_qr(1) = cmb_reftodoc.Text
        upddrequisitn_qr(2) = txt_year & "/" & CMB_DURILLMONTH
        upddrequisitn_qr(3) = Null
         
        upddrequisitn_qr(4) = Null
        upddrequisitn_qr(5) = Null
        If Len(Trim(txt_edsite.Text)) > 0 Then
            upddrequisitn_qr(6) = Null
            upddrequisitn_qr(6) = UCase(Trim(txt_edsite.Text))
        Else
            upddrequisitn_qr(6) = Null
        End If
        upddrequisitn_qr(7) = UCase(Trim(cmb_edprocedureheading.Text))
        upddrequisitn_qr(8) = Null
        upddrequisitn_qr(9) = Null
        TSSAMPLE = ""
        If list_edts.ListCount <> 0 Then
            For I = 0 To list_edts.ListCount - 1
                
                If list_edts.Selected(I) = True Then
                   TSSAMPLE = TSSAMPLE & list_edts.List(I) & " + "
                End If
            Next I
            TSSAMPLE = Left(TSSAMPLE, Len(Trim(TSSAMPLE)) - 1)
        End If
        upddrequisitn_qr(10) = Trim(TSSAMPLE)
        upddrequisitn_qr(11) = cmb_cnoslides.Text
        upddrequisitn_qr(12) = "NO"
        upddrequisitn_qr(13) = Null
        upddrequisitn_qr(14) = Null
        upddrequisitn_qr(15) = req_chglen(cmb_reqno.Text)
        upddrequisitn_qr.Execute
        '**** Updating the drequisition table
        
        
         '*** UPDATING THE PROVISIONAL DIAGNOSIS
        If opt_patient.Value = True Then
        If SELTA1012FLAG = False Then
            selta1012_qr(0) = Msk_caseno.FormattedText
            Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
            SELTA1012FLAG = True
        Else
             selta1012_qr(0) = Msk_caseno.FormattedText
             selta1012_rs.MoreResults
             selta1012_rs.Requery
        End If
        If selta1012_rs.RowCount > 0 Then
            UPDTA1012_QR(0) = Trim(txt_indctinvstg.Text)
            UPDTA1012_QR(1) = Msk_caseno.FormattedText
            UPDTA1012_QR.Execute
        Else
            insta1012_qr(0) = Msk_caseno.FormattedText
            insta1012_qr(1) = Trim(txt_indctinvstg.Text)
            insta1012_qr.Execute
        End If
        End If
        '*** UPDATING THE PROVISIONAL DIAGNOSIS
        '*** updating the clinical history
        If SELCLINHISTFLAG = False Then
            SELCLINHIST_QR(0) = Msk_caseno.FormattedText
            Set SELCLINHIST_RS = SELCLINHIST_QR.OpenResultset(1, 3)
            SELCLINHISTFLAG = True
        Else
            SELCLINHIST_QR(0) = Msk_caseno.FormattedText
            SELCLINHIST_RS.MoreResults
            SELCLINHIST_RS.Requery
        End If
        'If SELCLINHIST_RS.RowCount > 0 Then
'            updclinhist_qr(0) = Trim(txt_clinhist.Text)
'            updclinhist_qr(1) = Trim(Msk_caseno.FormattedText)
'            updclinhist_qr.Execute
'        Else
            insclinhist_qr(0) = Trim(Msk_caseno.FormattedText)
            insclinhist_qr(1) = GET_REVDATE(Trim(DATE11))
            insclinhist_qr(2) = Trim(time1)
            If Trim(txt_clinhist.Text) = "NYD" Then
                insclinhist_qr(3) = "NYD"
            Else
                insclinhist_qr(3) = Trim(txt_clinhist.Text)
            End If
            insclinhist_qr.Execute
        'End If
        
End Sub

Public Sub ECUPD()
'**** Updating the drequisition table
        upddrequisitn_qr(0) = (cmb_refdoc.Text)
        upddrequisitn_qr(1) = cmb_reftodoc.Text
        upddrequisitn_qr(2) = txt_year & "/" & CMB_DURILLMONTH
        upddrequisitn_qr(3) = Null
         
        upddrequisitn_qr(4) = Null
        upddrequisitn_qr(5) = Null
        upddrequisitn_qr(6) = Null
        upddrequisitn_qr(7) = Null
        upddrequisitn_qr(8) = Null
        upddrequisitn_qr(9) = Null
        upddrequisitn_qr(10) = Null
        upddrequisitn_qr(11) = Null
        upddrequisitn_qr(12) = "NO"
        upddrequisitn_qr(13) = Null
        upddrequisitn_qr(14) = Null
        upddrequisitn_qr(15) = req_chglen(cmb_reqno.Text)
        upddrequisitn_qr.Execute
        '**** Updating the drequisition table
        '**** UPDATING THE INSTRUCTION TABLE
'STR = "UPDATE DISLIB.INDICATION SET IND1=?,IND2,IND3=?,IND4=?,IND5=?,IND6=?,IND7=?,IND8=?,IND9=?,TP1=?,TP2=?,TP3=?,DRUG=?,DOSE=?,NOOFCYCLES=?,BSA=?,OTHERS=?, WHERE REQNO=?"
        If Len(Trim(ind1)) > 0 Then
            UPDINDICATIONQR(0) = Null
            UPDINDICATIONQR(0) = ind1
        Else
            UPDINDICATIONQR(0) = Null
        End If
        If Len(Trim(ind2)) > 0 Then
            UPDINDICATIONQR(1) = Null
            UPDINDICATIONQR(1) = ind2
        Else
            UPDINDICATIONQR(1) = Null
        End If
        If Len(Trim(ind3)) > 0 Then
            UPDINDICATIONQR(2) = Null
            UPDINDICATIONQR(2) = ind3
        Else
            UPDINDICATIONQR(2) = Null
        End If
        If Len(Trim(ind4)) > 0 Then
            UPDINDICATIONQR(3) = Null
            UPDINDICATIONQR(3) = ind4
        Else
            UPDINDICATIONQR(3) = Null
        End If
        If Len(Trim(ind5)) > 0 Then
            UPDINDICATIONQR(4) = Null
            UPDINDICATIONQR(4) = ind5
        Else
            UPDINDICATIONQR(4) = Null
        End If
        If Len(Trim(ind6)) > 0 Then
            UPDINDICATIONQR(5) = Null
            UPDINDICATIONQR(5) = ind6
        Else
            UPDINDICATIONQR(5) = Null
        End If
        If Len(Trim(ind7)) > 0 Then
            UPDINDICATIONQR(6) = Null
            UPDINDICATIONQR(6) = ind7
        Else
            UPDINDICATIONQR(6) = Null
        End If
        If Len(Trim(ind8)) > 0 Then
            UPDINDICATIONQR(7) = Null
            UPDINDICATIONQR(7) = ind8
        Else
            UPDINDICATIONQR(7) = Null
        End If
        If Len(Trim(ind9)) > 0 Then
            UPDINDICATIONQR(8) = Null
            UPDINDICATIONQR(8) = ind9
        Else
            UPDINDICATIONQR(8) = Null
        End If
        If Len(Trim(tp1)) > 0 Then
            UPDINDICATIONQR(9) = Null
            UPDINDICATIONQR(9) = tp1
        Else
            UPDINDICATIONQR(9) = Null
        End If
        If Len(Trim(tp2)) > 0 Then
            UPDINDICATIONQR(10) = Null
            UPDINDICATIONQR(10) = tp2
        Else
            UPDINDICATIONQR(10) = Null
        End If
        If Len(Trim(tp3)) > 0 Then
            UPDINDICATIONQR(11) = Null
            UPDINDICATIONQR(11) = tp3
        Else
            UPDINDICATIONQR(11) = Null
        End If
        'DRUG=?,DOSE=?,NOOFCYCLES=?,BSA=?,OTHERS=?, WHERE REQNO=?"
        UPDINDICATIONQR(12) = drug
        UPDINDICATIONQR(13) = dose
        UPDINDICATIONQR(14) = noofcycles
        UPDINDICATIONQR(15) = bsa
        If Len(Trim(OTHERS)) > 0 Then
            UPDINDICATIONQR(16) = Null
            UPDINDICATIONQR(16) = OTHERS
        Else
            UPDINDICATIONQR(16) = Null
        End If
        UPDINDICATIONQR(17) = req_chglen(cmb_reqno)
        UPDINDICATIONQR.Execute
         '*** UPDATING THE INSTRUCTION TABLE
         '*** UPDATING THE PROVISIONAL DIAGNOSIS
        If opt_patient.Value = True Then
        If SELTA1012FLAG = False Then
            selta1012_qr(0) = Msk_caseno.FormattedText
            Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
            SELTA1012FLAG = True
        Else
             selta1012_qr(0) = Msk_caseno.FormattedText
             selta1012_rs.MoreResults
             selta1012_rs.Requery
        End If
        If selta1012_rs.RowCount > 0 Then
            UPDTA1012_QR(0) = Trim(txt_indctinvstg.Text)
            UPDTA1012_QR(1) = Msk_caseno.FormattedText
            UPDTA1012_QR.Execute
        Else
            insta1012_qr(0) = Msk_caseno.FormattedText
            insta1012_qr(1) = Trim(txt_indctinvstg.Text)
            insta1012_qr.Execute
        End If
        End If
        '*** UPDATING THE PROVISIONAL DIAGNOSIS
        '*** updating the clinical history
        If SELCLINHISTFLAG = False Then
            SELCLINHIST_QR(0) = Msk_caseno.FormattedText
            Set SELCLINHIST_RS = SELCLINHIST_QR.OpenResultset(1, 3)
            SELCLINHISTFLAG = True
        Else
            SELCLINHIST_QR(0) = Msk_caseno.FormattedText
            SELCLINHIST_RS.MoreResults
            SELCLINHIST_RS.Requery
        End If
        'If SELCLINHIST_RS.RowCount > 0 Then
'            updclinhist_qr(0) = Trim(txt_clinhist.Text)
'            updclinhist_qr(1) = Trim(Msk_caseno.FormattedText)
'            updclinhist_qr.Execute
'        Else
            insclinhist_qr(0) = Trim(Msk_caseno.FormattedText)
            insclinhist_qr(1) = GET_REVDATE(Trim(DATE11))
            insclinhist_qr(2) = Trim(time1)
            If Trim(txt_clinhist.Text) = "NYD" Then
                insclinhist_qr(3) = "NYD"
            Else
                insclinhist_qr(3) = Trim(txt_clinhist.Text)
            End If
            insclinhist_qr.Execute
        'End If

End Sub

Public Sub SS()
    surgicalPathology.Visible = False
    immunohistochemistry.Visible = False
    frm_hematopathology.Visible = False
    biochemistry.Visible = False
    bceltm.Visible = True
    Label10.Visible = False
    Label11.Visible = False
    cmb_cyprocedure.Visible = False
    cmb_cyprocode.Visible = False
'    cardiology.Visible = False
    Microbiology.Visible = False
    CYTOLOGY.Visible = False
    fnac.Visible = False
    'cpl.Visible = False
    TUMOURMARKER.Visible = False
    hemcbc.Hide
    HEMBF.Hide
    HEMBMA.Hide
    dtmlab.show
    bloodProducts.Hide
    hlatyping.Hide
    lbl_req.Left = 3100
    dtmlab.Msk_caseno.Text = Msk_caseno.FormattedText
    dtmlab.txt_name.Text = txt_name.Text
    lbl_req.Caption = "REQUISITION - DTM LAB"
    pathology.Caption = "Pathology - DTM LAB"
    TEXT1.Visible = True
    TEXT1.Enabled = True
    Label1.Visible = True
'    If SELDTMNO_FLAG = False Then
'            SELDTMNO_QR(0) = Trim(msk_caseno.FormattedText)
'            Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
'            SELDTMNO_FLAG = True
'        Else
'           SELDTMNO_QR(0) = Trim(msk_caseno.FormattedText)
'           SELDTMNO_RS.MoreResults
'           SELDTMNO_RS.Requery
'        End If
'        If SELDTMNO_RS.RowCount > 0 Then
'            Text1.Text = Trim(SELDTMNO_RS("DTMNO"))
'            'bloodProducts.cmb_bg1.Text = Trim(SELDTMNO_RS("FLD1"))
'            Text1.Enabled = False
'            'bloodProducts.cmb_bg1.Enabled = False
'        Else
'            PATHNOAUTO
'        End If
        If SELDTMNO_FLAG = False Then
            SELDTMNO_QR(0) = Trim(Msk_caseno.FormattedText)
            Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
            SELDTMNO_FLAG = True
        Else
           SELDTMNO_QR(0) = Trim(Msk_caseno.FormattedText)
           SELDTMNO_RS.MoreResults
           SELDTMNO_RS.Requery
        End If
        If SELDTMNO_RS.RowCount > 0 Then
            TEXT1.Text = Trim(SELDTMNO_RS("DTMNO"))
            'cmb_bg1.Text = Trim(SELDTMNO_RS("FLD1"))
            'Text1.Enabled = False
            'bloodProducts.cmb_bg1.Enabled = False
        End If
End Sub

Public Sub BP()
    surgicalPathology.Visible = False
    immunohistochemistry.Visible = False
    frm_hematopathology.Visible = False
    biochemistry.Visible = False
    bceltm.Visible = True
    Label10.Visible = False
    Label11.Visible = False
    cmb_cyprocedure.Visible = False
    cmb_cyprocode.Visible = False
'    cardiology.Visible = False
    Microbiology.Visible = False
    CYTOLOGY.Visible = False
    fnac.Visible = False
    'cpl.Visible = False
    TUMOURMARKER.Visible = False
    hemcbc.Hide
    HEMBF.Hide
    HEMBMA.Hide
    dtmlab.Hide
    bloodProducts.show
    hlatyping.Hide
    lbl_req.Left = 3100
    bloodProducts.Msk_caseno.Text = Msk_caseno.FormattedText
    bloodProducts.txt_name.Text = txt_name.Text
    bloodProducts.TXT_DT.Text = Right(DATE11, 8)
    'bloodProducts.TXT_TM.Text = Trim(Left(time1, 2))
    
        If SELDTMNO_FLAG = False Then
           SELDTMNO_QR(0) = Trim(Msk_caseno.FormattedText)
           Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
           SELDTMNO_FLAG = True
        Else
           SELDTMNO_QR(0) = Trim(Msk_caseno.FormattedText)
           SELDTMNO_RS.MoreResults
           SELDTMNO_RS.Requery
        End If
        If SELDTMNO_RS.RowCount > 0 Then
            bloodProducts.txt_dtmno.Text = Trim(SELDTMNO_RS("DTMNO"))
            bloodProducts.cmb_bg1.Text = Trim(SELDTMNO_RS("FLD1"))
            bloodProducts.txt_dtmno.Enabled = False
            bloodProducts.cmb_bg1.Enabled = False
        End If
        
'        If SELDTMTAB_FLAG = False Then
'            SELDTMTAB_QR(0) = Trim(Msk_caseno.FormattedText)
'            Set SELDTMTAB_RS = SELDTMTAB_QR.OpenResultset(1, 3)
'            SELDTMTAB_FLAG = True
'        Else
'            SELDTMTAB_QR(0) = Trim(Msk_caseno.FormattedText)
'            SELDTMTAB_RS.MoreResults
'            SELDTMTAB_RS.Requery
'        End If
'        If SELDTMTAB_RS.RowCount > 0 Then
'            bloodProducts.cmb_bg1.Text = Trim(SELDTMTAB_RS("FLD1"))
'            bloodProducts.TXT_DTMNO.Enabled = False
'        End If
    lbl_req.Caption = "REQUISITION - BLOOD PRODUCTS"
    pathology.Caption = "Pathology - Blood Products"
    TEXT1.Visible = False
    Label1.Visible = False
    
End Sub

Public Sub HLASSIGN()
Dim durill As String
Dim I As Integer

On Error GoTo MSG:

''INSERT INTO DHEMCYTO TABLE
INSDHEMCYTO_QR(0) = req_chglen(Trim(txt_bcreqno.Text))
INSDHEMCYTO_QR(1) = MPO
INSDHEMCYTO_QR(2) = SBB
INSDHEMCYTO_QR(3) = PAS
INSDHEMCYTO_QR(4) = LAP
INSDHEMCYTO_QR(5) = NSE
INSDHEMCYTO_QR(6) = TRAP
INSDHEMCYTO_QR(7) = cbc
INSDHEMCYTO_QR(8) = hem
INSDHEMCYTO_QR(9) = cg1
INSDHEMCYTO_QR(10) = cyto
INSDHEMCYTO_QR(11) = smar
INSDHEMCYTO_QR(12) = TT
INSDHEMCYTO_QR(13) = oth
INSDHEMCYTO_QR(14) = othtext
INSDHEMCYTO_QR(15) = " "


INSDHEMCYTO_QR.Execute

'*** INSERTING INT0 RREQUISITION
INSRREQUISITN_QR(0) = req_chglen(Trim(txt_bcreqno.Text))
INSRREQUISITN_QR(1) = GET_REVDATE(CStr(DATE11))
INSRREQUISITN_QR(2) = Trim(txt_time.Text)
INSRREQUISITN_QR(3) = Trim(Msk_caseno.FormattedText)
INSRREQUISITN_QR(4) = Trim(txt_dept.Text)
INSRREQUISITN_QR(5) = Trim(txt_modcode.Text)
INSRREQUISITN_QR(6) = Trim(txt_status.Text)
INSRREQUISITN_QR(7) = Trim(txt_catgory.Text)
INSRREQUISITN_QR(8) = "8888"
INSRREQUISITN_QR(9) = UCase(Trim(cmb_refdoc.Text))
'INSRREQUISITN_QR(9) = Trim(txt_pridocser.Text)

'Added by Milind on 27-03-2006 for inserting date for the samples collected at outside
    If Check1.Value = 1 Then
        Dim Dt As String
        Dt = "18/02/1980"
        INSRREQUISITN_QR(15) = Format(CDate(Dt), "dd/mm/yyyy")
    Else
        INSRREQUISITN_QR(15) = Null
    End If
'upto here

    If SELRREQUISITNFLAG = False Then
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
        SELRREQUISITNFLAG = True
    Else
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        SELRREQUISITN_RS.MoreResults
        SELRREQUISITN_RS.Requery
    End If
    If SELRREQUISITN_RS.RowCount > 0 Then
        INSRREQUISITN_QR(10) = "YES"
    Else
        INSRREQUISITN_QR(10) = "NO"
    End If
If opt_patient.Value = True Then
    NO_CHARGE
    ''' Added by Sandhya dated 26.6.2009
    ''' for NC category & charged patient
    If txt_catgory.Text = "NC" And No_Chargeflag = True Then
        INSRREQUISITN_QR(11) = "NOCHARGE"
        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
        INSRREQUISITN_QR(13) = "P"
        'No_Chargeflag = False
    Else
        NOCHARG_QR(0) = Msk_caseno.FormattedText
        If NOCHARG_FLG = False Then
         Set NOCHARG_RS = NOCHARG_QR.OpenResultset(2, 3)
          NOCHARG_FLG = True
        Else
          NOCHARG_RS.MoreResults
          NOCHARG_RS.Requery
        End If
        If NOCHARG_RS.RowCount > 0 Then
           INSRREQUISITN_QR(11) = "NOCHARGE"
           INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
           INSRREQUISITN_QR(13) = "P"
        Else
            INSRREQUISITN_QR(11) = Null
            INSRREQUISITN_QR(12) = Null
            INSRREQUISITN_QR(13) = Null
        End If
       
    End If
End If
If opt_staff.Value = True Then
        INSRREQUISITN_QR(11) = "NOCHARGE"
        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
        INSRREQUISITN_QR(13) = "P"
End If
INSRREQUISITN_QR(14) = 0

INSRREQUISITN_QR.Execute
    '*** INSERTING INT0 RREQUISITION
    '*** INSERTING INTO DMULPROCE TABLE
    If CMB_TEXT.ListCount <> 0 Then
        For I = 0 To CMB_TEXT.ListCount - 1
            INSDMULPROCE_QR(0) = req_chglen(txt_bcreqno.Text)
            INSDMULPROCE_QR(1) = CMB_TEXT.List(I)
            INSDMULPROCE_QR(2) = GET_REVDATE(DATE11)
            INSDMULPROCE_QR(3) = Msk_caseno.FormattedText
            INSDMULPROCE_QR.Execute
        Next I
    End If
    
''*******check for zero amount***********
If Trim(txt_catgory.Text) = "C" Then
s_charges = ""
For I = 0 To CMB_TEXT.ListCount - 1
    If SELMA101hl_FLAG = False Then
        SELMA101hl_QR(0) = CMB_TEXT.List(I)
        Set SELMA101hl_RS = SELMA101hl_QR.OpenResultset(1, 3)
        SELMA101hl_FLAG = True
    Else
        SELMA101hl_QR(0) = CMB_TEXT.List(I)
        SELMA101hl_RS.MoreResults
        SELMA101hl_RS.Requery
    End If
    If SELMA101hl_RS("c_ch") > 0 Then
        s_charges = "Y"
    End If
Next I
If s_charges <> "Y" Then
    updreq_qr(0) = req_chglen(Trim(txt_bcreqno.Text))
    updreq_qr.Execute
End If
End If
'''''''''*********************************************

    '*** INSERTING INTO DMULPROCE TABLE
    '**** INSERTING INTO DREQUISITION

INSDREQUISITN_QR(0) = req_chglen(Trim(txt_bcreqno.Text))
INSDREQUISITN_QR(1) = " "
INSDREQUISITN_QR(2) = Trim(Msk_caseno.FormattedText)
INSDREQUISITN_QR(3) = Null

INSDREQUISITN_QR(4) = Null
dur_ill = txt_year & "/" & Trim(CMB_DURILLMONTH.Text)
INSDREQUISITN_QR(5) = dur_ill
INSDREQUISITN_QR(6) = bodyf
INSDREQUISITN_QR(7) = Null
INSDREQUISITN_QR(8) = Null
INSDREQUISITN_QR(9) = Null
INSDREQUISITN_QR(10) = Null
INSDREQUISITN_QR(11) = Null
INSDREQUISITN_QR(12) = Null
INSDREQUISITN_QR(13) = Null
INSDREQUISITN_QR(14) = Trim(cmb_refdoc.Text)
INSDREQUISITN_QR(15) = Trim(cmb_reftodoc.Text)
INSDREQUISITN_QR(16) = Null
If Trim(txt_status.Text) = "IP" Then
    INSDREQUISITN_QR(17) = Trim(txt_ward.Text)
    INSDREQUISITN_QR(18) = Trim(txt_bedno.Text)
Else
    INSDREQUISITN_QR(17) = Null
    INSDREQUISITN_QR(18) = Null
End If
INSDREQUISITN_QR(19) = ""
'INSDREQUISITN_QR(19) = DonorRelation
If Len(Trim(meterial)) <> 0 Then
    INSDREQUISITN_QR(20) = Null
    INSDREQUISITN_QR(20) = meterial
Else
    INSDREQUISITN_QR(20) = Null
    INSDREQUISITN_QR(20) = ""
End If
INSDREQUISITN_QR(21) = Null
INSDREQUISITN_QR.Execute
'**** INSERTING INTO DREQUISITION

 '*** UPDATING THE PROVISIONAL DIAGNOSIS
 
If opt_patient.Value = True Then
If SELTA1012FLAG = False Then
    selta1012_qr(0) = Msk_caseno.FormattedText
    Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
    SELTA1012FLAG = True
Else
     selta1012_qr(0) = Msk_caseno.FormattedText
     selta1012_rs.MoreResults
     selta1012_rs.Requery
End If
If selta1012_rs.RowCount > 0 Then
    UPDTA1012_QR(0) = Trim(txt_indctinvstg.Text)
    UPDTA1012_QR(1) = Msk_caseno.FormattedText
    UPDTA1012_QR.Execute
Else
    insta1012_qr(0) = Msk_caseno.FormattedText
    insta1012_qr(1) = Trim(txt_indctinvstg.Text)
    insta1012_qr.Execute
End If
End If
'*** UPDATING THE PROVISIONAL DIAGNOSIS
'*** INSERTING INTO DMULPROCE FOR COLLECTION AND DEPT COLLECTION SCREEN
''    INSDMULPROCE_QR(0) = req_chglen(txt_spreqno.Text)
''    INSDMULPROCE_QR(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
''    INSDMULPROCE_QR(2) = Msk_caseno.FormattedText
''    INSDMULPROCE_QR(3) = cmb_procedurecode.Text
''    INSDMULPROCE_QR.Execute

'***INSERTING INTO DMULPROCE FOR COLLECTION AND DEPT COLLECTION SCREEN
'***INSERTING INTO DREQSTAFFD TABLE
If opt_staff.Value = True Then
    INSDREQSTAFFD_QR(0) = req_chglen(txt_bcreqno.Text)
    INSDREQSTAFFD_QR(1) = Msk_caseno.Text
    INSDREQSTAFFD_QR(2) = Trim(txt_name.Text)
    INSDREQSTAFFD_QR(3) = Trim(txt_Age)
    INSDREQSTAFFD_QR(4) = Trim(txt_sex)
    If Len(Trim(cmb_dependents.Text)) = 0 Then
        INSDREQSTAFFD_QR(5) = "SELF"
    Else
        I = InStr(cmb_dependents, "$")
        INSDREQSTAFFD_QR(5) = Mid(cmb_dependents, I + 1, 10)
    End If
    INSDREQSTAFFD_QR.Execute
End If
'***INSERTING INTO DREQSTAFFD TABLE
    If SELCLINHISTFLAG = False Then
        SELCLINHIST_QR(0) = Msk_caseno.FormattedText
        Set SELCLINHIST_RS = SELCLINHIST_QR.OpenResultset(1, 3)
        SELCLINHISTFLAG = True
    Else
        SELCLINHIST_QR(0) = Msk_caseno.FormattedText
        SELCLINHIST_RS.MoreResults
        SELCLINHIST_RS.Requery
    End If
'    If SELCLINHIST_RS.RowCount > 0 Then
'        updclinhist_qr(0) = Trim(txt_clinhist.Text)
'        updclinhist_qr(1) = Trim(Msk_caseno.Text)
'        updclinhist_qr.Execute
'    Else
        insclinhist_qr(0) = Trim(Msk_caseno.FormattedText)
        insclinhist_qr(1) = GET_REVDATE(Trim(DATE11))
        insclinhist_qr(2) = Trim(time1)
        If Trim(txt_clinhist.Text) = "NYD" Then
            insclinhist_qr(3) = "NYD"
        Else
            insclinhist_qr(3) = Trim(txt_clinhist.Text)
        End If
        insclinhist_qr.Execute
    'End If

'***** UPDATE DREQUISITN TABLE Added by Shraddha Pandit on 09/06/09 *****
If SELDISICU_FLAG = False Then
    SELDISICU_QR(0) = Trim(ipadd)
    Set SELDISICU_RS = SELDISICU_QR.OpenResultset(1, 3)
    SELDISICU_FLAG = True
Else
    SELDISICU_QR(0) = Trim(ipadd)
    SELDISICU_RS.MoreResults
    SELDISICU_RS.Requery
End If

If SELDISICU_RS.RowCount > 0 Then
    If Trim(SELDISICU_RS(1)) = "10.3.1.56" Then
        UPDICU_QR(0) = "YR"
        UPDICU_QR(1) = req_chglen(Trim(txt_bcreqno.Text))
        UPDICU_QR.Execute
    Else
        UPDICU_QR(0) = "YES"
        UPDICU_QR(1) = req_chglen(Trim(txt_bcreqno.Text))
        UPDICU_QR.Execute
    End If
End If

'*************************************************************************



Data1.UpdateRecord
Data1.Recordset.Bookmark = Data1.Recordset.LastModified
Exit Sub
MSG:
MsgBox Err.Description

End Sub

Public Sub BMASSIGN()
Dim durill As String
Dim I As Integer

On Error GoTo MSG:
'*** INSERTING INT0 RREQUISITION
INSRREQUISITN_QR(0) = req_chglen(Trim(txt_bcreqno.Text))
INSRREQUISITN_QR(1) = GET_REVDATE(CStr(DATE11))
INSRREQUISITN_QR(2) = Trim(txt_time.Text)
INSRREQUISITN_QR(3) = Trim(Msk_caseno.FormattedText)
INSRREQUISITN_QR(4) = Trim(txt_dept.Text)
INSRREQUISITN_QR(5) = Trim(txt_modcode.Text)
INSRREQUISITN_QR(6) = Trim(txt_status.Text)
INSRREQUISITN_QR(7) = Trim(txt_catgory.Text)
INSRREQUISITN_QR(8) = "8888"
INSRREQUISITN_QR(9) = UCase(Trim(cmb_refdoc.Text))
'INSRREQUISITN_QR(9) = Trim(txt_pridocser.Text)

'Added by Milind on 27-03-2006 for inserting date for the samples collected at outside
    If Check1.Value = 1 Then
        Dim Dt As String
        Dt = "18/02/1980"
        INSRREQUISITN_QR(15) = Format(CDate(Dt), "dd/mm/yyyy")
    Else
        INSRREQUISITN_QR(15) = Null
    End If
'upto here

    If SELRREQUISITNFLAG = False Then
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
        SELRREQUISITNFLAG = True
    Else
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        SELRREQUISITN_RS.MoreResults
        SELRREQUISITN_RS.Requery
    End If
    If SELRREQUISITN_RS.RowCount > 0 Then
        INSRREQUISITN_QR(10) = "YES"
    Else
        INSRREQUISITN_QR(10) = "NO"
    End If
If opt_patient.Value = True Then
    NO_CHARGE
    ''' Added by Sandhya dated 26.6.2009
    ''' for NC category & charged patient
    If txt_catgory.Text = "NC" And No_Chargeflag = True Then
        INSRREQUISITN_QR(11) = "NOCHARGE"
        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
        INSRREQUISITN_QR(13) = "P"
        'No_Chargeflag = False
    Else
        
            INSRREQUISITN_QR(11) = Null
            INSRREQUISITN_QR(12) = Null
            INSRREQUISITN_QR(13) = Null
       
    End If
End If
If opt_staff.Value = True Then
        INSRREQUISITN_QR(11) = "NOCHARGE"
        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
        INSRREQUISITN_QR(13) = "P"
End If
INSRREQUISITN_QR(14) = 0

INSRREQUISITN_QR.Execute
    '*** INSERTING INT0 RREQUISITION
    '*** INSERTING INTO DMULPROCE TABLE
    If CMB_TEXT.ListCount <> 0 Then
        For I = 0 To CMB_TEXT.ListCount - 1
            INSDMULPROCE_QR(0) = req_chglen(txt_bcreqno.Text)
            INSDMULPROCE_QR(1) = CMB_TEXT.List(I)
            INSDMULPROCE_QR(2) = GET_REVDATE(DATE11)
            INSDMULPROCE_QR(3) = Msk_caseno.FormattedText
            INSDMULPROCE_QR.Execute
        Next I
    End If
    '*** INSERTING INTO DMULPROCE TABLE
    ''*******check for zero amount***********
 If Trim(txt_catgory.Text) = "C" Then
    s_charges = ""
    For I = 0 To CMB_TEXT.ListCount - 1
        If SELMA101hl_FLAG = False Then
            SELMA101hl_QR(0) = CMB_TEXT.List(I)
            Set SELMA101hl_RS = SELMA101hl_QR.OpenResultset(1, 3)
            SELMA101hl_FLAG = True
        Else
            SELMA101hl_QR(0) = CMB_TEXT.List(I)
            SELMA101hl_RS.MoreResults
            SELMA101hl_RS.Requery
        End If
        If SELMA101hl_RS("c_ch") > 0 Then
            s_charges = "Y"
        End If
    Next I
    If s_charges <> "Y" Then
        updreq_qr(0) = req_chglen(Trim(txt_bcreqno.Text))
        updreq_qr.Execute
    End If
  End If
    '''''''''*********************************************





    '*** INSERTING INTO DBMACIND
        INSDBMAINDQR(0) = req_chglen(txt_bcreqno.Text)
        If Len(Trim(ind1)) > 0 Then
            INSDBMAINDQR(1) = ind1
        Else
            INSDBMAINDQR(1) = Null
        End If
        If Len(Trim(ind2)) > 0 Then
            INSDBMAINDQR(2) = ind2
        Else
            INSDBMAINDQR(2) = Null
        End If
        If Len(Trim(ind3)) > 0 Then
            INSDBMAINDQR(3) = ind3
        Else
            INSDBMAINDQR(3) = Null
        End If
        If Len(Trim(ind4)) > 0 Then
            INSDBMAINDQR(4) = ind4
        Else
            INSDBMAINDQR(4) = Null
        End If
        If Len(Trim(ind5)) > 0 Then
            INSDBMAINDQR(5) = ind5
        Else
            INSDBMAINDQR(5) = Null
        End If
        If Len(Trim(ind6)) > 0 Then
            INSDBMAINDQR(6) = ind6
        Else
            INSDBMAINDQR(6) = Null
        End If
        If Len(Trim(ind7)) > 0 Then
            INSDBMAINDQR(7) = ind7
        Else
            INSDBMAINDQR(7) = Null
        End If
        If Len(Trim(ind8)) > 0 Then
            INSDBMAINDQR(8) = ind8
        Else
            INSDBMAINDQR(8) = Null
        End If
        If Len(Trim(ind9)) > 0 Then
            INSDBMAINDQR(9) = ind9
        Else
            INSDBMAINDQR(9) = Null
        End If
        'INSDBMAINDQR(9) = Null
        INSDBMAINDQR(10) = Null
        If Len(Trim(tp1)) > 0 Then
            INSDBMAINDQR(11) = tp1
        Else
            INSDBMAINDQR(11) = Null
        End If
        If Len(Trim(tp2)) > 0 Then
            INSDBMAINDQR(12) = tp2
        Else
            INSDBMAINDQR(12) = Null
        End If
        If Len(Trim(tp3)) > 0 Then
            INSDBMAINDQR(13) = tp3
        Else
            INSDBMAINDQR(13) = Null
        End If
        INSDBMAINDQR.Execute
        
    
    '*** INSERTING INTO DBMACIND
    
    
    '**** INSERTING INTO DREQUISITION

INSDREQUISITN_QR(0) = req_chglen(Trim(txt_bcreqno.Text))
INSDREQUISITN_QR(1) = " "
INSDREQUISITN_QR(2) = Trim(Msk_caseno.FormattedText)
INSDREQUISITN_QR(3) = Null

INSDREQUISITN_QR(4) = Null
dur_ill = txt_year & "/" & Trim(CMB_DURILLMONTH.Text)
INSDREQUISITN_QR(5) = dur_ill
INSDREQUISITN_QR(6) = bodyf
INSDREQUISITN_QR(7) = Null
INSDREQUISITN_QR(8) = Null
INSDREQUISITN_QR(9) = Null
INSDREQUISITN_QR(10) = Null
INSDREQUISITN_QR(11) = Null
INSDREQUISITN_QR(12) = Null
INSDREQUISITN_QR(13) = Null
INSDREQUISITN_QR(14) = Trim(cmb_refdoc.Text)
INSDREQUISITN_QR(15) = Trim(cmb_reftodoc.Text)
INSDREQUISITN_QR(16) = Null
If Trim(txt_status.Text) = "IP" Then
    INSDREQUISITN_QR(17) = Trim(txt_ward.Text)
    INSDREQUISITN_QR(18) = Trim(txt_bedno.Text)
Else
    INSDREQUISITN_QR(17) = Null
    INSDREQUISITN_QR(18) = Null
End If
INSDREQUISITN_QR(19) = " "
If Len(Trim(meterial)) <> 0 Then
    INSDREQUISITN_QR(20) = Null
    INSDREQUISITN_QR(20) = meterial
Else
    INSDREQUISITN_QR(20) = Null
    INSDREQUISITN_QR(20) = ""
End If
INSDREQUISITN_QR(21) = Null
INSDREQUISITN_QR.Execute
'**** INSERTING INTO DREQUISITION

 '*** UPDATING THE PROVISIONAL DIAGNOSIS
 
If opt_patient.Value = True Then
If SELTA1012FLAG = False Then
    selta1012_qr(0) = Msk_caseno.FormattedText
    Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
    SELTA1012FLAG = True
Else
     selta1012_qr(0) = Msk_caseno.FormattedText
     selta1012_rs.MoreResults
     selta1012_rs.Requery
End If
If selta1012_rs.RowCount > 0 Then
    UPDTA1012_QR(0) = Trim(txt_indctinvstg.Text)
    UPDTA1012_QR(1) = Msk_caseno.FormattedText
    UPDTA1012_QR.Execute
Else
    insta1012_qr(0) = Msk_caseno.FormattedText
    insta1012_qr(1) = Trim(txt_indctinvstg.Text)
    insta1012_qr.Execute
End If
End If
'*** UPDATING THE PROVISIONAL DIAGNOSIS
'*** INSERTING INTO DMULPROCE FOR COLLECTION AND DEPT COLLECTION SCREEN
''    INSDMULPROCE_QR(0) = req_chglen(txt_spreqno.Text)
''    INSDMULPROCE_QR(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
''    INSDMULPROCE_QR(2) = Msk_caseno.FormattedText
''    INSDMULPROCE_QR(3) = cmb_procedurecode.Text
''    INSDMULPROCE_QR.Execute

'***INSERTING INTO DMULPROCE FOR COLLECTION AND DEPT COLLECTION SCREEN
'***INSERTING INTO DREQSTAFFD TABLE
If opt_staff.Value = True Then
    INSDREQSTAFFD_QR(0) = req_chglen(txt_bcreqno.Text)
    INSDREQSTAFFD_QR(1) = Msk_caseno.Text
    INSDREQSTAFFD_QR(2) = Trim(txt_name.Text)
    INSDREQSTAFFD_QR(3) = Trim(txt_Age)
    INSDREQSTAFFD_QR(4) = Trim(txt_sex)
    If Len(Trim(cmb_dependents.Text)) = 0 Then
        INSDREQSTAFFD_QR(5) = "SELF"
    Else
        I = InStr(cmb_dependents, "$")
        INSDREQSTAFFD_QR(5) = Mid(cmb_dependents, I + 1, 10)
    End If
    INSDREQSTAFFD_QR.Execute
End If
'***INSERTING INTO DREQSTAFFD TABLE
    If SELCLINHISTFLAG = False Then
        SELCLINHIST_QR(0) = Msk_caseno.FormattedText
        Set SELCLINHIST_RS = SELCLINHIST_QR.OpenResultset(1, 3)
        SELCLINHISTFLAG = True
    Else
        SELCLINHIST_QR(0) = Msk_caseno.FormattedText
        SELCLINHIST_RS.MoreResults
        SELCLINHIST_RS.Requery
    End If
'    If SELCLINHIST_RS.RowCount > 0 Then
'        updclinhist_qr(0) = Trim(txt_clinhist.Text)
'        updclinhist_qr(1) = Trim(Msk_caseno.Text)
'        updclinhist_qr.Execute
'    Else
        insclinhist_qr(0) = Trim(Msk_caseno.FormattedText)
        insclinhist_qr(1) = GET_REVDATE(Trim(DATE11))
        insclinhist_qr(2) = Trim(time1)
        If Trim(txt_clinhist.Text) = "NYD" Then
            insclinhist_qr(3) = "NYD"
        Else
            insclinhist_qr(3) = Trim(txt_clinhist.Text)
        End If
        insclinhist_qr.Execute
    'End If

Data1.UpdateRecord
Data1.Recordset.Bookmark = Data1.Recordset.LastModified
Exit Sub
MSG:
MsgBox Err.Description
rdb.RollbackTrans
End Sub

Public Sub BFASSIGN()
Dim durill As String
Dim I As Integer

On Error GoTo MSG:
'*** INSERTING INT0 RREQUISITION
INSRREQUISITN_QR(0) = req_chglen(Trim(txt_bcreqno.Text))
INSRREQUISITN_QR(1) = GET_REVDATE(CStr(DATE11))
INSRREQUISITN_QR(2) = Trim(txt_time.Text)
INSRREQUISITN_QR(3) = Trim(Msk_caseno.FormattedText)
INSRREQUISITN_QR(4) = Trim(txt_dept.Text)
INSRREQUISITN_QR(5) = Trim(txt_modcode.Text)
INSRREQUISITN_QR(6) = Trim(txt_status.Text)
INSRREQUISITN_QR(7) = Trim(txt_catgory.Text)
INSRREQUISITN_QR(8) = "8888"
INSRREQUISITN_QR(9) = UCase(Trim(cmb_refdoc.Text))
'INSRREQUISITN_QR(9) = Trim(txt_pridocser.Text)

'Added by Milind on 27-03-2006 for inserting date for the samples collected at outside
    If Check1.Value = 1 Then
        Dim Dt As String
        Dt = "18/02/1980"
        INSRREQUISITN_QR(15) = Format(CDate(Dt), "dd/mm/yyyy")
    Else
        INSRREQUISITN_QR(15) = Null
    End If
'upto here

    If SELRREQUISITNFLAG = False Then
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
        SELRREQUISITNFLAG = True
    Else
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        SELRREQUISITN_RS.MoreResults
        SELRREQUISITN_RS.Requery
    End If
    If SELRREQUISITN_RS.RowCount > 0 Then
        INSRREQUISITN_QR(10) = "YES"
    Else
        INSRREQUISITN_QR(10) = "NO"
    End If
If opt_patient.Value = True Then
    NO_CHARGEFA
    ''' Added by Sandhya dated 26.6.2009
    ''' for NC category & charged patient
    If txt_catgory.Text = "NC" And No_Chargeflag = True Then
        INSRREQUISITN_QR(11) = "NOCHARGE"
        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
        INSRREQUISITN_QR(13) = "P"
        'No_Chargeflag = False
    Else
        
            INSRREQUISITN_QR(11) = Null
            INSRREQUISITN_QR(12) = Null
            INSRREQUISITN_QR(13) = Null
       
    End If
End If
If opt_staff.Value = True Then
        INSRREQUISITN_QR(11) = "NOCHARGE"
        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
        INSRREQUISITN_QR(13) = "P"
End If
INSRREQUISITN_QR(14) = 0

INSRREQUISITN_QR.Execute
    '*** INSERTING INT0 RREQUISITION
    '*** INSERTING INTO DMULPROCE TABLE
    If CMB_TEXT.ListCount <> 0 Then
        For I = 0 To CMB_TEXT.ListCount - 1
            INSDMULPROCE_QR(0) = req_chglen(txt_bcreqno.Text)
            INSDMULPROCE_QR(1) = CMB_TEXT.List(I)
            INSDMULPROCE_QR(2) = GET_REVDATE(DATE11)
            INSDMULPROCE_QR(3) = Trim(Msk_caseno.FormattedText)
            INSDMULPROCE_QR.Execute
        Next I
    End If
    
    ''*******check for zero amount***********
  If Trim(txt_catgory.Text) = "C" Then
    s_charges = ""
    For I = 0 To CMB_TEXT.ListCount - 1
        If SELMA101hl_FLAG = False Then
            SELMA101hl_QR(0) = CMB_TEXT.List(I)
            Set SELMA101hl_RS = SELMA101hl_QR.OpenResultset(1, 3)
            SELMA101hl_FLAG = True
        Else
            SELMA101hl_QR(0) = CMB_TEXT.List(I)
            SELMA101hl_RS.MoreResults
            SELMA101hl_RS.Requery
        End If
        If SELMA101hl_RS("c_ch") > 0 Then
            s_charges = "Y"
        End If
    Next I
    If s_charges <> "Y" Then
        updreq_qr(0) = req_chglen(Trim(txt_bcreqno.Text))
        updreq_qr.Execute
    End If
  End If
    '''''''''*********************************************

    
    '*** INSERTING INTO DMULPROCE TABLE
    
    '*** INSERTING INTO DBMACIND
        INSDBMAINDQR(0) = req_chglen(txt_bcreqno.Text)
        If Len(Trim(ind1)) > 0 Then
            INSDBMAINDQR(1) = ind1
        Else
            INSDBMAINDQR(1) = Null
        End If
        If Len(Trim(ind2)) > 0 Then
            INSDBMAINDQR(2) = ind2
        Else
            INSDBMAINDQR(2) = Null
        End If
        If Len(Trim(ind3)) > 0 Then
            INSDBMAINDQR(3) = ind3
        Else
            INSDBMAINDQR(3) = Null
        End If
        If Len(Trim(ind4)) > 0 Then
            INSDBMAINDQR(4) = ind4
        Else
            INSDBMAINDQR(4) = Null
        End If
        If Len(Trim(ind5)) > 0 Then
            INSDBMAINDQR(5) = ind5
        Else
            INSDBMAINDQR(5) = Null
        End If
        'If Len(Trim(ind6)) > 0 Then
        '   INSDBMAINDQR(6) = ind6
        'Else
            INSDBMAINDQR(6) = Null
        'End If
        'If Len(Trim(ind7)) > 0 Then
         '   INSDBMAINDQR(7) = ind7
        'Else
            INSDBMAINDQR(7) = Null
        'End If
        'If Len(Trim(ind8)) > 0 Then
        '    INSDBMAINDQR(8) = ind8
        'Else
            INSDBMAINDQR(8) = Null
        'End If
        INSDBMAINDQR(9) = Null
        INSDBMAINDQR(10) = Null
        INSDBMAINDQR(11) = Null
        INSDBMAINDQR(12) = Null
        INSDBMAINDQR(13) = Null
        INSDBMAINDQR.Execute
        
    
    '*** INSERTING INTO DBMACIND
    
    
    
    '**** INSERTING INTO DREQUISITION

INSDREQUISITN_QR(0) = req_chglen(Trim(txt_bcreqno.Text))
INSDREQUISITN_QR(1) = " "
INSDREQUISITN_QR(2) = Trim(Msk_caseno.FormattedText)
INSDREQUISITN_QR(3) = Null

INSDREQUISITN_QR(4) = Null
dur_ill = txt_year & "/" & Trim(CMB_DURILLMONTH.Text)
INSDREQUISITN_QR(5) = dur_ill
INSDREQUISITN_QR(6) = bodyf
INSDREQUISITN_QR(7) = Null
INSDREQUISITN_QR(8) = Null
INSDREQUISITN_QR(9) = Null
INSDREQUISITN_QR(10) = Null
INSDREQUISITN_QR(11) = Null
INSDREQUISITN_QR(12) = Null
INSDREQUISITN_QR(13) = Null
INSDREQUISITN_QR(14) = Trim(cmb_refdoc.Text)
INSDREQUISITN_QR(15) = Trim(cmb_reftodoc.Text)
INSDREQUISITN_QR(16) = Null
If Trim(txt_status.Text) = "IP" Then
    INSDREQUISITN_QR(17) = Trim(txt_ward.Text)
    INSDREQUISITN_QR(18) = Trim(txt_bedno.Text)
Else
    INSDREQUISITN_QR(17) = Null
    INSDREQUISITN_QR(18) = Null
End If
INSDREQUISITN_QR(19) = " "
If Len(Trim(meterial)) <> 0 Then
    INSDREQUISITN_QR(20) = Null
    INSDREQUISITN_QR(20) = meterial
Else
    INSDREQUISITN_QR(20) = Null
    INSDREQUISITN_QR(20) = ""
End If
INSDREQUISITN_QR(21) = Null
INSDREQUISITN_QR.Execute
'**** INSERTING INTO DREQUISITION

 '*** UPDATING THE PROVISIONAL DIAGNOSIS
 
If opt_patient.Value = True Then
If SELTA1012FLAG = False Then
    selta1012_qr(0) = Msk_caseno.FormattedText
    Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
    SELTA1012FLAG = True
Else
     selta1012_qr(0) = Msk_caseno.FormattedText
     selta1012_rs.MoreResults
     selta1012_rs.Requery
End If
If selta1012_rs.RowCount > 0 Then
    UPDTA1012_QR(0) = Trim(txt_indctinvstg.Text)
    UPDTA1012_QR(1) = Msk_caseno.FormattedText
    UPDTA1012_QR.Execute
Else
    insta1012_qr(0) = Msk_caseno.FormattedText
    insta1012_qr(1) = Trim(txt_indctinvstg.Text)
    insta1012_qr.Execute
End If
End If
'*** UPDATING THE PROVISIONAL DIAGNOSIS
'*** INSERTING INTO DMULPROCE FOR COLLECTION AND DEPT COLLECTION SCREEN
''    INSDMULPROCE_QR(0) = req_chglen(txt_spreqno.Text)
''    INSDMULPROCE_QR(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
''    INSDMULPROCE_QR(2) = Msk_caseno.FormattedText
''    INSDMULPROCE_QR(3) = cmb_procedurecode.Text
''    INSDMULPROCE_QR.Execute

'***INSERTING INTO DMULPROCE FOR COLLECTION AND DEPT COLLECTION SCREEN
'***INSERTING INTO DREQSTAFFD TABLE
If opt_staff.Value = True Then
    INSDREQSTAFFD_QR(0) = req_chglen(txt_bcreqno.Text)
    INSDREQSTAFFD_QR(1) = Msk_caseno.Text
    INSDREQSTAFFD_QR(2) = Trim(txt_name.Text)
    INSDREQSTAFFD_QR(3) = Trim(txt_Age)
    INSDREQSTAFFD_QR(4) = Trim(txt_sex)
    If Len(Trim(cmb_dependents.Text)) = 0 Then
        INSDREQSTAFFD_QR(5) = "SELF"
    Else
        I = InStr(cmb_dependents, "$")
        INSDREQSTAFFD_QR(5) = Mid(cmb_dependents, I + 1, 10)
    End If
    INSDREQSTAFFD_QR.Execute
End If
'***INSERTING INTO DREQSTAFFD TABLE
    If SELCLINHISTFLAG = False Then
        SELCLINHIST_QR(0) = Msk_caseno.FormattedText
        Set SELCLINHIST_RS = SELCLINHIST_QR.OpenResultset(1, 3)
        SELCLINHISTFLAG = True
    Else
        SELCLINHIST_QR(0) = Msk_caseno.FormattedText
        SELCLINHIST_RS.MoreResults
        SELCLINHIST_RS.Requery
    End If
'    If SELCLINHIST_RS.RowCount > 0 Then
'        updclinhist_qr(0) = Trim(txt_clinhist.Text)
'        updclinhist_qr(1) = Trim(Msk_caseno.Text)
'        updclinhist_qr.Execute
'    Else
        insclinhist_qr(0) = Trim(Msk_caseno.FormattedText)
        insclinhist_qr(1) = GET_REVDATE(Trim(DATE11))
        insclinhist_qr(2) = Trim(time1)
        If Trim(txt_clinhist.Text) = "NYD" Then
            insclinhist_qr(3) = "NYD"
        Else
            insclinhist_qr(3) = Trim(txt_clinhist.Text)
        End If
        insclinhist_qr.Execute
    'End If

Data1.UpdateRecord
Data1.Recordset.Bookmark = Data1.Recordset.LastModified
Exit Sub
MSG:
MsgBox Err.Description
rdb.RollbackTrans
End Sub

Public Sub HLVIEW()
'If SELDMULPROCERS("PRO_CODE") = "U723" Then
 '   hemcbc.CHK_MORPHOLOGYDLC.Value = 1
'End If
If SELDMULPROCERS("PRO_CODE") = "U723" Then
    hemcbc.CHK_HAEMOGRAM.Value = 1
End If
'If SELDMULPROCERS("PRO_CODE") = "U722" Then
'    hemcbc.CHK_DIFFCOUNT.Value = 1
'End If
'If SELDMULPROCERS("PRO_CODE") = "U721" Then
'    hemcbc.CHK_HAEMOGRAM.Value = 1
'End If
If SELDMULPROCERS("PRO_CODE") = "U706" Then
    hemcbc.CHK_ESR.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "U713" Then
    hemcbc.CHK_MALARIALPARASITE.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "U724" Then
    hemcbc.CHK_RETICCOUNT.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "U753" Then
    hemcbc.CHK_SMC.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "F754" Then
    hemcbc.chk_sm.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "U752" Then
    hemcbc.chk_cytochemistry.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "U712" Then
    hemcbc.CHK_FIBRINOGEN.Value = 1
    hemcbc.CHK_CP.Value = 1
    hemcbc.chk_pt.Value = 1
    hemcbc.CHK_PTTK.Value = 1
    hemcbc.CHK_FDPDDIMER.Value = 1
    hemcbc.CHK_FIBRINOGEN.Value = 1
    
End If
If SELDMULPROCERS("PRO_CODE") = "U711" Then
    hemcbc.CHK_CP.Value = 1
    hemcbc.chk_pt.Value = 1
    hemcbc.CHK_PTTK.Value = 1
    
End If
If SELDMULPROCERS("PRO_CODE") = "F710" Then
   'hemcbc.chk_bt.Value = 1
    hemcbc.CHK_PTTK.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "U708" Then
    hemcbc.chk_pt.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "F707" Then
    hemcbc.CHK_BT.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "F709" Then
    hemcbc.CHK_PTTK.Value = 1
    hemcbc.chk_pt.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "U714" Then
    hemcbc.CHK_FDPDDIMER.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "U715" Then
    hemcbc.CHK_FIBRINOGEN.Value = 1
End If
End Sub

Public Sub BMVIEW()
If SELDMULPROCERS("PRO_CODE") = "U751" Then
    HEMBMA.CHK_MORPHOLOGY.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "U752" Then
    HEMBMA.chk_cytochemistry.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "U754" Then
    HEMBMA.chk_sm.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "U753" Then
    HEMBMA.CHK_SMC.Value = 1
End If
End Sub

Public Sub BFVIEW()
' SUNDEEP CHECK
If SELDMULPROCERS("PRO_CODE") = "U718" Then
    HEMBF.chk_csf.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "U716" Then
    HEMBF.CHK_UPT.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "U727" Then
    HEMBF.chk_pfa.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "U726" Then
    HEMBF.CHK_PF.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "U725" Then
    HEMBF.chk_af.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "U717" Then
    HEMBF.chk_sa.Value = 1
End If
' SINGLE SURFACE MARKER
If SELDMULPROCERS("PRO_CODE") = "U754" Then
    HEMBF.chk_sm.Value = 1
End If
' COMPLETE SURFACE MARKER
If SELDMULPROCERS("PRO_CODE") = "U753" Then
    HEMBF.CHK_SMC.Value = 1
End If

End Sub

Public Sub HT()
 surgicalPathology.Visible = False
    immunohistochemistry.Visible = False
    frm_hematopathology.Visible = False
    biochemistry.Visible = False
    bceltm.Visible = True
    Label10.Visible = False
    Label11.Visible = False
    cmb_cyprocedure.Visible = False
    cmb_cyprocode.Visible = False
'    cardiology.Visible = False
    Microbiology.Visible = False
    CYTOLOGY.Visible = False
    fnac.Visible = False
    'cpl.Visible = False
    TUMOURMARKER.Visible = False
    hemcbc.Hide
    HEMBF.Hide
    HEMBMA.Hide
    dtmlab.Hide
    bloodProducts.Hide
    hlatyping.show
    
    hlatyping.Msk_caseno.Text = Msk_caseno.FormattedText
    hlatyping.txt_name.Text = txt_name.Text
    lbl_req.Left = 3100
    lbl_req.Caption = "REQUISITION - HLA TYPING"
    pathology.Caption = "Transfusion Medicine - HLA TYPING"
    TEXT1.Visible = False
    Label1.Visible = False
'    If SELDTMNO_FLAG = False Then
'            SELDTMNO_QR(0) = Trim(msk_caseno.FormattedText)
'            Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
'            SELDTMNO_FLAG = True
'        Else
'           SELDTMNO_QR(0) = Trim(msk_caseno.FormattedText)
'           SELDTMNO_RS.MoreResults
'           SELDTMNO_RS.Requery
'        End If
'        If SELDTMNO_RS.RowCount > 0 Then
'            'bloodProducts.TXT_DTMNO.Text = Trim(SELDTMNO_RS("DTMNO"))
'            'bloodProducts.cmb_bg1.Text = Trim(SELDTMNO_RS("FLD1"))
'            'bloodProducts.TXT_DTMNO.Enabled = False
'            'bloodProducts.cmb_bg1.Enabled = False
'            Text1.Text = Trim(SELDTMNO_RS("DTMNO"))
'            Text1.Enabled = False
'        Else
'            PATHNOAUTO
'        End If
End Sub
Public Sub BPASSIGN()
Dim durill As String
Dim I As Integer

'On Error GoTo MSG:
'*** INSERTING INT0 RREQUISITION
INSRREQUISITN_QR(0) = req_chglen(Trim(txt_bcreqno.Text))
INSRREQUISITN_QR(1) = GET_REVDATE(CStr(DATE11))
INSRREQUISITN_QR(2) = Trim(txt_time.Text)
INSRREQUISITN_QR(3) = Trim(Msk_caseno.FormattedText)
INSRREQUISITN_QR(4) = Trim(txt_dept.Text)
INSRREQUISITN_QR(5) = Trim(txt_modcode.Text)
INSRREQUISITN_QR(6) = Trim(txt_status.Text)
INSRREQUISITN_QR(7) = Trim(txt_catgory.Text)
INSRREQUISITN_QR(8) = "8888"
INSRREQUISITN_QR(9) = UCase(Trim(cmb_refdoc.Text))
'INSRREQUISITN_QR(9) = Trim(txt_pridocser.Text)

'Added by Milind on 27-03-2006 for inserting date for the samples collected at outside
    If Check1.Value = 1 Then
        Dim Dt As String
        Dt = "18/02/1980"
        INSRREQUISITN_QR(15) = Format(CDate(Dt), "dd/mm/yyyy")
    Else
        INSRREQUISITN_QR(15) = Null
    End If
'upto here

    If SELRREQUISITNFLAG = False Then
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
        SELRREQUISITNFLAG = True
    Else
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        SELRREQUISITN_RS.MoreResults
        SELRREQUISITN_RS.Requery
    End If
    If SELRREQUISITN_RS.RowCount > 0 Then
        INSRREQUISITN_QR(10) = "YES"
    Else
        INSRREQUISITN_QR(10) = "NO"
    End If
If opt_patient.Value = True Then
    NO_CHARGE
    ''' Added by Sandhya dated 26.6.2009
    ''' for NC category & charged patient
    If txt_catgory.Text = "NC" And No_Chargeflag = True Then
        INSRREQUISITN_QR(11) = "NOCHARGE"
        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
        INSRREQUISITN_QR(13) = "P"
        'No_Chargeflag = False
        ''ICU
    ElseIf lbl_icureq.Visible = True Then
        'If Trim(CMB_ICU1.Text) = "YES" Then
            INSRREQUISITN_QR(11) = Null
            INSRREQUISITN_QR(12) = Null
            'INSRREQUISITN_QR(13) = "P"
            'INSRREQUISITN_QR(13) = " "
             INSRREQUISITN_QR(13) = Null
        'End If
    ''ICU
    Else
            ' suggested for RDP,SDP,FFP,FVIIID,CRYO,ALB ON 14-01-04
        If (((Trim(Len(ind3)) > 0 Or Trim(Len(ind4)) > 0 Or Trim(Len(ind5)) > 0 Or Trim(Len(ind6)) > 0 Or Trim(Len(ind7)) > 0 Or Trim(Len(ind8)) > 0)) And ((Trim(Len(ind1)) = 0 And Trim(Len(ind2)) = 0))) Then
            INSRREQUISITN_QR(11) = "NOCHARGE"
            INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
            INSRREQUISITN_QR(13) = "P"
        Else
            ' suggested for RDP,SDP,FFP,FVIIID,CRYO,ALB ON 14-01-04
            INSRREQUISITN_QR(11) = Null
            INSRREQUISITN_QR(12) = Null
            INSRREQUISITN_QR(13) = "P" 'Null 'COMMENT
                ' suggested for RDP,SDP,FFP,FVIIID,CRYO,ALB ON 14-01-04
        End If
                ' suggested for RDP,SDP,FFP,FVIIID,CRYO,ALB ON 14-01-04
    End If
End If
If opt_staff.Value = True Then
        INSRREQUISITN_QR(11) = "NOCHARGE"
        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
        INSRREQUISITN_QR(13) = "P"
End If
INSRREQUISITN_QR(14) = 0

INSRREQUISITN_QR.Execute
    '*** INSERTING INT0 RREQUISITION
    '*** INSERTING INTO DMULPROCE TABLE
    Dim H002 As String
    If CMB_TEXT.ListCount <> 0 Then
        For I = 0 To CMB_TEXT.ListCount - 1
            INSDMULPROCE_QR(0) = req_chglen(txt_bcreqno.Text)
            INSDMULPROCE_QR(1) = CMB_TEXT.List(I)
            INSDMULPROCE_QR(2) = GET_REVDATE(DATE11)
            INSDMULPROCE_QR(3) = Msk_caseno.FormattedText
            INSDMULPROCE_QR.Execute
            If CMB_TEXT.List(I) = "H002" And (txt_catgory.Text <> "NC") Then
              H002 = "YES"
            End If
        Next I
    End If
    '*** INSERTING INTO DMULPROCE TABLE
    If H002 = "YES" Then
      UPDREQNSTAT_QR(0) = req_chglen(txt_bcreqno.Text)
      UPDREQNSTAT_QR.Execute
      'SUNDEEP DR.NARAYAN

      'SUNDEEP DR.NARAYAN
      UPDDMULPROCE_QR(0) = crossmatch
      UPDDMULPROCE_QR(1) = req_chglen(txt_bcreqno.Text)
      UPDDMULPROCE_QR.Execute

    End If

    ''************FOR BP(RDP,SDP,FFP,FVIIID,CRYO,ALB) UPDATE MEMO NO AS NOCHARGE **************
    If (((Trim(Len(ind3)) > 0 Or Trim(Len(ind4)) > 0 Or Trim(Len(ind5)) > 0 Or Trim(Len(ind6)) > 0 Or Trim(Len(ind7)) > 0 Or Trim(Len(ind8)) > 0)) And ((Trim(Len(ind1)) = 0 And Trim(Len(ind2)) = 0))) Then
        UPDMEMORREQ_QR(0) = "NOCHARGE"
        UPDMEMORREQ_QR(1) = GET_REVDATE(DATE11)
        UPDMEMORREQ_QR(2) = req_chglen(txt_bcreqno.Text)
        UPDMEMORREQ_QR(3) = Msk_caseno.FormattedText
        UPDMEMORREQ_QR.Execute
    End If
    '************FOR BP(RDP,SDP,FFP,FVIIID,CRYO,ALB) UPDATE MEMO NO AS NOCHARGE **************
    '*** INSERTING INTO BLOOD PRODUCT
        INSBLOODPRODUQR(0) = req_chglen(txt_bcreqno.Text)

        If Len(Trim(BG)) > 0 Then
            INSBLOODPRODUQR(1) = Trim(BG)
        Else
            INSBLOODPRODUQR(1) = Null
        End If

        If Len(Trim(RFT)) > 0 Then
            INSBLOODPRODUQR(2) = Trim(RFT)
        Else
            INSBLOODPRODUQR(2) = Null
        End If
        If Len(Trim(HB)) > 0 Then
           INSBLOODPRODUQR(3) = Trim(HB)
        Else
            INSBLOODPRODUQR(3) = Null

        End If
        If Len(Trim(REQUION)) > 0 Then
            INSBLOODPRODUQR(4) = GET_REVDATE(REQUION)
        Else
            INSBLOODPRODUQR(4) = Null
        End If
        If Len(Trim(T_M)) > 0 And Len(Trim(T_M1)) > 0 Then
            INSBLOODPRODUQR(5) = Trim(T_M) & Trim(T_M1)
        Else
            INSBLOODPRODUQR(5) = Null
        End If
        If Len(Trim(TYPEOFREQ)) > 0 Then
            INSBLOODPRODUQR(6) = TYPEOFREQ
        Else
            INSBLOODPRODUQR(6) = Null
        End If

        If Len(Trim(ind1)) > 0 Then
            INSBLOODPRODUQR(7) = ind1
        Else
            INSBLOODPRODUQR(7) = Null
        End If
        If Len(Trim(ind2)) > 0 Then
            INSBLOODPRODUQR(8) = ind2
        Else
            INSBLOODPRODUQR(8) = Null
        End If
        If Len(Trim(ind3)) > 0 Then
            INSBLOODPRODUQR(9) = ind3
        Else
            INSBLOODPRODUQR(9) = Null
        End If
        If Len(Trim(ind4)) > 0 Then
            INSBLOODPRODUQR(10) = ind4
        Else
            INSBLOODPRODUQR(10) = Null
        End If
        If Len(Trim(ind5)) > 0 Then
            INSBLOODPRODUQR(11) = ind5
        Else
            INSBLOODPRODUQR(11) = Null
        End If
        If Len(Trim(ind6)) > 0 Then
            INSBLOODPRODUQR(12) = ind6
        Else
            INSBLOODPRODUQR(12) = Null
        End If
        If Len(Trim(ind7)) > 0 Then
            INSBLOODPRODUQR(13) = ind7
        Else
            INSBLOODPRODUQR(13) = Null
        End If
        If Len(Trim(ind8)) > 0 Then
            INSBLOODPRODUQR(14) = ind8
        Else
            INSBLOODPRODUQR(14) = Null
        End If
        If Len(Trim(SR1)) > 0 Then
            INSBLOODPRODUQR(15) = SR1
        Else
            INSBLOODPRODUQR(15) = Null
        End If
        If Len(Trim(SR2)) > 0 Then
            INSBLOODPRODUQR(16) = SR2
        Else
            INSBLOODPRODUQR(16) = Null
        End If
        ' sundeep Pinpol TBPRODUCT
            If Trim(SR1) <> "" And Trim(SR2) <> "" Then
                IL = "IL"
            End If
            If Trim(SR1) <> "" And Trim(SR2) = "" Then
                IL = "I"
            End If
            If Trim(SR1) = "" And Trim(SR2) <> "" Then
                IL = "L"
            End If
        '' sundeep Pinpol TBPRODUCT
        If Len(Trim(OTHERS1)) > 0 Then
            INSBLOODPRODUQR(17) = OTHERS1
        Else
            INSBLOODPRODUQR(17) = Null
        End If
        If Len(Trim(PHOBOTMIST)) > 0 Then
            INSBLOODPRODUQR(18) = PHOBOTMIST
        Else
            INSBLOODPRODUQR(18) = Null
        End If
        If Len(Trim(NOOFUNREP)) > 0 Then
            INSBLOODPRODUQR(19) = NOOFUNREP
        Else
            INSBLOODPRODUQR(19) = Null
        End If
        If Len(Trim(T_M1)) > 0 Then
            INSBLOODPRODUQR(20) = T_M1
        Else
            INSBLOODPRODUQR(20) = Null
        End If
        If Len(Trim(FLD)) > 0 Then
            INSBLOODPRODUQR(21) = Trim(FLD)
        Else
            INSBLOODPRODUQR(21) = Null
        End If
        If Len(Trim(PLTC)) > 0 Then
            INSBLOODPRODUQR(22) = Trim(PLTC)
        Else
            INSBLOODPRODUQR(22) = Null
        End If
        INSBLOODPRODUQR(23) = wbunit
        INSBLOODPRODUQR(24) = pcunit
        INSBLOODPRODUQR.Execute


    '*** INSERTING INTO BLOOD PRODUCT
    '*** inserting into TBPRODUCT of Pimpol
        'FCaseNo , FREQNNO, FCOMPONENT, FQTY, FREMARKS, FSPLREQ, FPHLEB, FREPLCDNO, FISSUED, FFLAG
'''        'WB

' THIS CODE HIDED FOR DISPLY POINT VALUES IN BP REPORT BY  04/06/2004
'///////////////////////////////////////////////////////////////////////////
        If Trim(bpwb) <> "" Then
        Ins_TBProduct_qr(0) = Trim(Msk_caseno.FormattedText)
        Ins_TBProduct_qr(1) = Trim(req_chglen(txt_bcreqno.Text))
        Ins_TBProduct_qr(2) = Trim(bpwb)
        Ins_TBProduct_qr(3) = ind1
        Ins_TBProduct_qr(4) = Trim(OTHERS1)
        Ins_TBProduct_qr(5) = Trim(IL)
        Ins_TBProduct_qr(6) = Trim(PHOBOTMIST)
        Ins_TBProduct_qr(7) = 0
        Ins_TBProduct_qr(8) = 0
        Ins_TBProduct_qr(9) = Null
        Ins_TBProduct_qr.Execute
        End If
        'PC
        If Trim(bppc) <> "" Then
        Ins_TBProduct_qr(0) = Trim(Msk_caseno.FormattedText)
        Ins_TBProduct_qr(1) = Trim(req_chglen(txt_bcreqno.Text))
        Ins_TBProduct_qr(2) = Trim(bppc)
        Ins_TBProduct_qr(3) = ind2
        Ins_TBProduct_qr(4) = Trim(OTHERS1)
        Ins_TBProduct_qr(5) = Trim(IL)
        Ins_TBProduct_qr(6) = Trim(PHOBOTMIST)
        Ins_TBProduct_qr(7) = 0
        Ins_TBProduct_qr(8) = 0
        Ins_TBProduct_qr(9) = Null
        Ins_TBProduct_qr.Execute
        End If
        'RDP
        If Trim(bprdp) <> "" Then
        Ins_TBProduct_qr(0) = Trim(Msk_caseno.FormattedText)
        Ins_TBProduct_qr(1) = Trim(req_chglen(txt_bcreqno.Text))
        Ins_TBProduct_qr(2) = Trim(bprdp)
        Ins_TBProduct_qr(3) = ind3
        Ins_TBProduct_qr(4) = Trim(OTHERS1)
        Ins_TBProduct_qr(5) = Trim(IL)
        Ins_TBProduct_qr(6) = Trim(PHOBOTMIST)
        Ins_TBProduct_qr(7) = 0
        Ins_TBProduct_qr(8) = 0
        Ins_TBProduct_qr(9) = Null
        Ins_TBProduct_qr.Execute
        End If

         'sdp
        If Trim(bpsdp) <> "" Then
        Ins_TBProduct_qr(0) = Trim(Msk_caseno.FormattedText)
        Ins_TBProduct_qr(1) = Trim(req_chglen(txt_bcreqno.Text))
        Ins_TBProduct_qr(2) = Trim(bpsdp)
        Ins_TBProduct_qr(3) = ind4
        Ins_TBProduct_qr(4) = Trim(OTHERS1)
        Ins_TBProduct_qr(5) = Trim(IL)
        Ins_TBProduct_qr(6) = Trim(PHOBOTMIST)
        Ins_TBProduct_qr(7) = 0
        Ins_TBProduct_qr(8) = 0
        Ins_TBProduct_qr(9) = Null
        Ins_TBProduct_qr.Execute
        End If
        'ffp
        If Trim(bpffp) <> "" Then
        Ins_TBProduct_qr(0) = Trim(Msk_caseno.FormattedText)
        Ins_TBProduct_qr(1) = Trim(req_chglen(txt_bcreqno.Text))
        Ins_TBProduct_qr(2) = Trim(bpffp)
        Ins_TBProduct_qr(3) = ind5
        Ins_TBProduct_qr(4) = Trim(OTHERS1)
        Ins_TBProduct_qr(5) = Trim(IL)
        Ins_TBProduct_qr(6) = Trim(PHOBOTMIST)
        Ins_TBProduct_qr(7) = 0
        Ins_TBProduct_qr(8) = 0
        Ins_TBProduct_qr(9) = Null
        Ins_TBProduct_qr.Execute
        End If
        'FVIIID
        If Trim(bpfviiid) <> "" Then
        Ins_TBProduct_qr(0) = Trim(Msk_caseno.FormattedText)
        Ins_TBProduct_qr(1) = Trim(req_chglen(txt_bcreqno.Text))
        Ins_TBProduct_qr(2) = Trim(bpfviiid)
        Ins_TBProduct_qr(3) = ind6
        Ins_TBProduct_qr(4) = Trim(OTHERS1)
        Ins_TBProduct_qr(5) = Trim(IL)
        Ins_TBProduct_qr(6) = Trim(PHOBOTMIST)
        Ins_TBProduct_qr(7) = 0
        Ins_TBProduct_qr(8) = 0
        Ins_TBProduct_qr(9) = Null
        Ins_TBProduct_qr.Execute
        End If

        'CRYO
        If Trim(bpcryo) <> "" Then
        Ins_TBProduct_qr(0) = Trim(Msk_caseno.FormattedText)
        Ins_TBProduct_qr(1) = Trim(req_chglen(txt_bcreqno.Text))
        Ins_TBProduct_qr(2) = Trim(bpcryo)
        Ins_TBProduct_qr(3) = ind7
        Ins_TBProduct_qr(4) = Trim(OTHERS1)
        Ins_TBProduct_qr(5) = Trim(IL)
        Ins_TBProduct_qr(6) = Trim(PHOBOTMIST)
        Ins_TBProduct_qr(7) = 0
        Ins_TBProduct_qr(8) = 0
        Ins_TBProduct_qr(9) = Null
        Ins_TBProduct_qr.Execute
        End If

        'alb
        If Trim(bpalb) <> "" Then
        Ins_TBProduct_qr(0) = Trim(Msk_caseno.FormattedText)
        Ins_TBProduct_qr(1) = Trim(req_chglen(txt_bcreqno.Text))
        Ins_TBProduct_qr(2) = Trim(bpalb)
        Ins_TBProduct_qr(3) = ind8
        Ins_TBProduct_qr(4) = Trim(OTHERS1)
        Ins_TBProduct_qr(5) = Trim(IL)
        Ins_TBProduct_qr(6) = Trim(PHOBOTMIST)
        Ins_TBProduct_qr(7) = 0
        Ins_TBProduct_qr(8) = 0
        Ins_TBProduct_qr(9) = Null
        Ins_TBProduct_qr.Execute
        End If

'///////////////////////////////////////////////////////////////////////
'THAT CODE HIDED BY  FOR DISPLAY POINT VALUE IN BP REPORT(.7)04/06/2004

    '*** inserting into TBPRODUCT of Pimpol

    '**** INSERTING INTO DREQUISITION


INSDREQUISITN_QR(0) = req_chglen(Trim(txt_bcreqno.Text))
If DTMNO = "" Then
    INSDREQUISITN_QR(1) = " "
Else
    INSDREQUISITN_QR(1) = Trim(DTMNO)
End If
INSDREQUISITN_QR(2) = Trim(Msk_caseno.FormattedText)
'INSDREQUISITN_QR(3) = Null
'Commented and added by Shraddha Pandit on 06/03/09
INSDREQUISITN_QR(3) = aliquot

INSDREQUISITN_QR(4) = Null
dur_ill = txt_year & "/" & Trim(CMB_DURILLMONTH.Text)
INSDREQUISITN_QR(5) = dur_ill
INSDREQUISITN_QR(6) = bodyf
INSDREQUISITN_QR(7) = Null
INSDREQUISITN_QR(8) = Null
INSDREQUISITN_QR(9) = Null
INSDREQUISITN_QR(10) = Null
INSDREQUISITN_QR(11) = Null
INSDREQUISITN_QR(12) = Null
INSDREQUISITN_QR(13) = Null
INSDREQUISITN_QR(14) = Trim(cmb_refdoc.Text)
INSDREQUISITN_QR(15) = Trim(cmb_reftodoc.Text)
INSDREQUISITN_QR(16) = Null
If Trim(txt_status.Text) = "IP" Then
    INSDREQUISITN_QR(17) = Trim(txt_ward.Text)
    INSDREQUISITN_QR(18) = Trim(txt_bedno.Text)
Else
    INSDREQUISITN_QR(17) = Null
    INSDREQUISITN_QR(18) = Null
End If
INSDREQUISITN_QR(19) = " "
If Len(Trim(meterial)) <> 0 Then
    INSDREQUISITN_QR(20) = Null
    INSDREQUISITN_QR(20) = meterial
Else
    INSDREQUISITN_QR(20) = Null
    INSDREQUISITN_QR(20) = ""
End If
INSDREQUISITN_QR(21) = Null
INSDREQUISITN_QR.Execute
'**** INSERTING INTO DREQUISITION

 '*** UPDATING THE PROVISIONAL DIAGNOSIS

If opt_patient.Value = True Then
    If SELTA1012FLAG = False Then
        selta1012_qr(0) = Msk_caseno.FormattedText
        Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
        SELTA1012FLAG = True
    Else
         selta1012_qr(0) = Msk_caseno.FormattedText
         selta1012_rs.MoreResults
         selta1012_rs.Requery
    End If
    If selta1012_rs.RowCount > 0 Then
        UPDTA1012_QR(0) = Trim(txt_indctinvstg.Text)
        UPDTA1012_QR(1) = Msk_caseno.FormattedText
        UPDTA1012_QR.Execute
    Else
        insta1012_qr(0) = Msk_caseno.FormattedText
        insta1012_qr(1) = Trim(txt_indctinvstg.Text)
        insta1012_qr.Execute
    End If
End If
'*** UPDATING THE PROVISIONAL DIAGNOSIS
'*** INSERTING INTO DMULPROCE FOR COLLECTION AND DEPT COLLECTION SCREEN
''    INSDMULPROCE_QR(0) = req_chglen(txt_spreqno.Text)
''    INSDMULPROCE_QR(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
''    INSDMULPROCE_QR(2) = Msk_caseno.FormattedText
''    INSDMULPROCE_QR(3) = cmb_procedurecode.Text
''    INSDMULPROCE_QR.Execute

'***INSERTING INTO DMULPROCE FOR COLLECTION AND DEPT COLLECTION SCREEN
'***INSERTING INTO DREQSTAFFD TABLE
If opt_staff.Value = True Then
    INSDREQSTAFFD_QR(0) = req_chglen(txt_bcreqno.Text)
    INSDREQSTAFFD_QR(1) = Msk_caseno.Text
    INSDREQSTAFFD_QR(2) = Trim(txt_name.Text)
    INSDREQSTAFFD_QR(3) = Trim(txt_Age)
    INSDREQSTAFFD_QR(4) = Trim(txt_sex)
    If Len(Trim(cmb_dependents.Text)) = 0 Then
        INSDREQSTAFFD_QR(5) = "SELF"
    Else
        I = InStr(cmb_dependents, "$")
        INSDREQSTAFFD_QR(5) = Mid(cmb_dependents, I + 1, 10)
    End If
    INSDREQSTAFFD_QR.Execute
End If
'***INSERTING INTO DREQSTAFFD TABLE
    If SELCLINHISTFLAG = False Then
        SELCLINHIST_QR(0) = Msk_caseno.FormattedText
        Set SELCLINHIST_RS = SELCLINHIST_QR.OpenResultset(1, 3)
        SELCLINHISTFLAG = True
    Else
        SELCLINHIST_QR(0) = Msk_caseno.FormattedText
        SELCLINHIST_RS.MoreResults
        SELCLINHIST_RS.Requery
    End If
'    If SELCLINHIST_RS.RowCount > 0 Then
'        updclinhist_qr(0) = Trim(txt_clinhist.Text)
'        updclinhist_qr(1) = Trim(Msk_caseno.Text)
'        updclinhist_qr.Execute
'    Else
        insclinhist_qr(0) = Trim(Msk_caseno.FormattedText)
        insclinhist_qr(1) = GET_REVDATE(Trim(DATE11))
        insclinhist_qr(2) = Trim(time1)
        If Trim(txt_clinhist.Text) = "NYD" Then
            insclinhist_qr(3) = "NYD"
        Else
            insclinhist_qr(3) = Trim(txt_clinhist.Text)
        End If
        insclinhist_qr.Execute
    'End If
''INSERT DTMNO INTO DTMNO TABLE IF THEY INSERT DTMNO INTO DTMNO FIELD.

        If SELDTMNO_FLAG = False Then
            SELDTMNO_QR(0) = Trim(Msk_caseno.FormattedText)
            Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
            SELDTMNO_FLAG = True
        Else
           SELDTMNO_QR(0) = Trim(Msk_caseno.FormattedText)
           SELDTMNO_RS.MoreResults
           SELDTMNO_RS.Requery
        End If
        If SELDTMNO_RS.RowCount = 0 Then
            INSDTMNO_QR(0) = Trim(Msk_caseno.FormattedText)
            INSDTMNO_QR(1) = Trim(DTMNO)
            INSDTMNO_QR(2) = Trim(BG)
            INSDTMNO_QR.Execute
        End If

''UPDATE DREQUISITN TABLE
If SELDISICU_FLAG = False Then
    SELDISICU_QR(0) = Trim(ipadd)
    Set SELDISICU_RS = SELDISICU_QR.OpenResultset(1, 3)
    SELDISICU_FLAG = True
Else
    SELDISICU_QR(0) = Trim(ipadd)
    SELDISICU_RS.MoreResults
    SELDISICU_RS.Requery
End If
If Not SELDISICU_RS.EOF Then
    UPDICU_QR(0) = "YES"
    UPDICU_QR(1) = req_chglen(Trim(txt_bcreqno.Text))
    UPDICU_QR.Execute
End If

''UPDATE DREQUISITN TABLE
Data1.UpdateRecord
Data1.Recordset.Bookmark = Data1.Recordset.LastModified
Exit Sub
'MSG:
'MsgBox Err.Description
rdb.RollbackTrans
cmd_cancel_Click
End Sub

Public Sub SSASSIGN()
Dim durill As String
Dim I As Integer

On Error GoTo MSG:
'*** INSERTING INT0 RREQUISITION
INSRREQUISITN_QR(0) = req_chglen(Trim(txt_bcreqno.Text))
INSRREQUISITN_QR(1) = GET_REVDATE(CStr(DATE11))
INSRREQUISITN_QR(2) = Trim(txt_time.Text)
INSRREQUISITN_QR(3) = Trim(Msk_caseno.FormattedText)
INSRREQUISITN_QR(4) = Trim(txt_dept.Text)
INSRREQUISITN_QR(5) = Trim(txt_modcode.Text)
INSRREQUISITN_QR(6) = Trim(txt_status.Text)
INSRREQUISITN_QR(7) = Trim(txt_catgory.Text)
INSRREQUISITN_QR(8) = "8888"
INSRREQUISITN_QR(9) = UCase(Trim(cmb_refdoc.Text))
'INSRREQUISITN_QR(9) = Trim(txt_pridocser.Text)

'Added by Milind on 27-03-2006 for inserting date for the samples collected at outside
    If Check1.Value = 1 Then
        Dim Dt As String
        Dt = "18/02/1980"
        INSRREQUISITN_QR(15) = Format(CDate(Dt), "dd/mm/yyyy")
    Else
        INSRREQUISITN_QR(15) = Null
    End If
'upto here

    If SELRREQUISITNFLAG = False Then
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
        SELRREQUISITNFLAG = True
    Else
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        SELRREQUISITN_RS.MoreResults
        SELRREQUISITN_RS.Requery
    End If
    If SELRREQUISITN_RS.RowCount > 0 Then
        INSRREQUISITN_QR(10) = "YES"
    Else
        INSRREQUISITN_QR(10) = "NO"
    End If
If opt_patient.Value = True Then
    NO_CHARGE
    ''' Added by Sandhya dated 26.6.2009
    ''' for NC category & charged patient
    If txt_catgory.Text = "NC" And No_Chargeflag = True Then
        INSRREQUISITN_QR(11) = "NOCHARGE"
        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
        INSRREQUISITN_QR(13) = "P"
        'No_Chargeflag = False
    Else
        
            INSRREQUISITN_QR(11) = Null
            INSRREQUISITN_QR(12) = Null
            INSRREQUISITN_QR(13) = Null
       
    End If
End If
If opt_staff.Value = True Then
        INSRREQUISITN_QR(11) = "NOCHARGE"
        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
        INSRREQUISITN_QR(13) = "P"
End If
INSRREQUISITN_QR(14) = 0

INSRREQUISITN_QR.Execute
    '*** INSERTING INT0 RREQUISITION
    '*** INSERTING INTO DMULPROCE TABLE
    If CMB_TEXT.ListCount <> 0 Then
        For I = 0 To CMB_TEXT.ListCount - 1
            INSDMULPROCE_QR(0) = req_chglen(txt_bcreqno.Text)
            INSDMULPROCE_QR(1) = CMB_TEXT.List(I)
            INSDMULPROCE_QR(2) = GET_REVDATE(DATE11)
            INSDMULPROCE_QR(3) = Msk_caseno.FormattedText
            INSDMULPROCE_QR.Execute
        Next I
    End If
    '*** INSERTING INTO DMULPROCE TABLE
    
    '**** INSERTING INTO DREQUISITION

INSDREQUISITN_QR(0) = req_chglen(Trim(txt_bcreqno.Text))
If TEXT1.Text = "" Then
    INSDREQUISITN_QR(1) = " "
Else
    INSDREQUISITN_QR(1) = Trim(TEXT1.Text)
End If
INSDREQUISITN_QR(2) = Trim(Msk_caseno.FormattedText)
INSDREQUISITN_QR(3) = Null

INSDREQUISITN_QR(4) = Null
dur_ill = txt_year & "/" & Trim(CMB_DURILLMONTH.Text)
INSDREQUISITN_QR(5) = dur_ill
INSDREQUISITN_QR(6) = bodyf
INSDREQUISITN_QR(7) = Null
INSDREQUISITN_QR(8) = Null
INSDREQUISITN_QR(9) = Null
INSDREQUISITN_QR(10) = Null
INSDREQUISITN_QR(11) = Null
INSDREQUISITN_QR(12) = Null
INSDREQUISITN_QR(13) = Null
INSDREQUISITN_QR(14) = Trim(cmb_refdoc.Text)
INSDREQUISITN_QR(15) = Trim(cmb_reftodoc.Text)
INSDREQUISITN_QR(16) = Null
INSDREQUISITN_QR(16) = HTOther
If Trim(txt_status.Text) = "IP" Then
    INSDREQUISITN_QR(17) = Trim(txt_ward.Text)
    INSDREQUISITN_QR(18) = Trim(txt_bedno.Text)
Else
    INSDREQUISITN_QR(17) = Null
    INSDREQUISITN_QR(18) = Null
End If
INSDREQUISITN_QR(19) = " "
If Len(Trim(meterial)) <> 0 Then
    INSDREQUISITN_QR(20) = Null
    INSDREQUISITN_QR(20) = meterial
Else
    INSDREQUISITN_QR(20) = Null
    INSDREQUISITN_QR(20) = ""
End If
INSDREQUISITN_QR(21) = Null
INSDREQUISITN_QR.Execute
'**** INSERTING INTO DREQUISITION

 '*** UPDATING THE PROVISIONAL DIAGNOSIS
 
If opt_patient.Value = True Then
If SELTA1012FLAG = False Then
    selta1012_qr(0) = Msk_caseno.FormattedText
    Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
    SELTA1012FLAG = True
Else
     selta1012_qr(0) = Msk_caseno.FormattedText
     selta1012_rs.MoreResults
     selta1012_rs.Requery
End If
If selta1012_rs.RowCount > 0 Then
    UPDTA1012_QR(0) = Trim(txt_indctinvstg.Text)
    UPDTA1012_QR(1) = Msk_caseno.FormattedText
    UPDTA1012_QR.Execute
Else
    insta1012_qr(0) = Msk_caseno.FormattedText
    insta1012_qr(1) = Trim(txt_indctinvstg.Text)
    insta1012_qr.Execute
End If
End If
'*** UPDATING THE PROVISIONAL DIAGNOSIS
'*** INSERTING INTO DMULPROCE FOR COLLECTION AND DEPT COLLECTION SCREEN
''    INSDMULPROCE_QR(0) = req_chglen(txt_spreqno.Text)
''    INSDMULPROCE_QR(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
''    INSDMULPROCE_QR(2) = Msk_caseno.FormattedText
''    INSDMULPROCE_QR(3) = cmb_procedurecode.Text
''    INSDMULPROCE_QR.Execute

'***INSERTING INTO DMULPROCE FOR COLLECTION AND DEPT COLLECTION SCREEN
'***INSERTING INTO DREQSTAFFD TABLE
If opt_staff.Value = True Then
    INSDREQSTAFFD_QR(0) = req_chglen(txt_bcreqno.Text)
    INSDREQSTAFFD_QR(1) = Msk_caseno.Text
    INSDREQSTAFFD_QR(2) = Trim(txt_name.Text)
    INSDREQSTAFFD_QR(3) = Trim(txt_Age)
    INSDREQSTAFFD_QR(4) = Trim(txt_sex)
    If Len(Trim(cmb_dependents.Text)) = 0 Then
        INSDREQSTAFFD_QR(5) = "SELF"
    Else
        I = InStr(cmb_dependents, "$")
        INSDREQSTAFFD_QR(5) = Mid(cmb_dependents, I + 1, 10)
    End If
    INSDREQSTAFFD_QR.Execute
End If
'***INSERTING INTO DREQSTAFFD TABLE
    If SELCLINHISTFLAG = False Then
        SELCLINHIST_QR(0) = Msk_caseno.FormattedText
        Set SELCLINHIST_RS = SELCLINHIST_QR.OpenResultset(1, 3)
        SELCLINHISTFLAG = True
    Else
        SELCLINHIST_QR(0) = Msk_caseno.FormattedText
        SELCLINHIST_RS.MoreResults
        SELCLINHIST_RS.Requery
    End If
'    If SELCLINHIST_RS.RowCount > 0 Then
'        updclinhist_qr(0) = Trim(txt_clinhist.Text)
'        updclinhist_qr(1) = Trim(Msk_caseno.Text)
'        updclinhist_qr.Execute
'    Else
        insclinhist_qr(0) = Trim(Msk_caseno.FormattedText)
        insclinhist_qr(1) = GET_REVDATE(Trim(DATE11))
        insclinhist_qr(2) = Trim(time1)
        If Trim(txt_clinhist.Text) = "NYD" Then
            insclinhist_qr(3) = "NYD"
        Else
            insclinhist_qr(3) = Trim(txt_clinhist.Text)
        End If
        insclinhist_qr.Execute
    'End If
If Trim(BG1) = "Y" Then
    If SELDTMNO_FLAG = False Then
        SELDTMNO_QR(0) = Trim(Msk_caseno.FormattedText)
        Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
        SELDTMNO_FLAG = True
    Else
        SELDTMNO_QR(0) = Trim(Msk_caseno.FormattedText)
        SELDTMNO_RS.MoreResults
        SELDTMNO_RS.Requery
    End If
    If SELDTMNO_RS.RowCount = 0 Then
        INSDTMNO_QR(0) = Trim(Msk_caseno.FormattedText)
        INSDTMNO_QR(1) = Trim(TEXT1.Text)
        INSDTMNO_QR(2) = ""
        INSDTMNO_QR.Execute
    Else
        PATHNOAUTO
    End If
End If
Data1.UpdateRecord
Data1.Recordset.Bookmark = Data1.Recordset.LastModified
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Public Sub HTASSIGN()
Dim durill As String
Dim I As Integer

On Error GoTo MSG:
'*** INSERTING INT0 RREQUISITION
INSRREQUISITN_QR(0) = req_chglen(Trim(txt_bcreqno.Text))
INSRREQUISITN_QR(1) = GET_REVDATE(CStr(DATE11))
INSRREQUISITN_QR(2) = Trim(txt_time.Text)
INSRREQUISITN_QR(3) = Trim(Msk_caseno.FormattedText)
INSRREQUISITN_QR(4) = Trim(txt_dept.Text)
INSRREQUISITN_QR(5) = Trim(txt_modcode.Text)
INSRREQUISITN_QR(6) = Trim(txt_status.Text)
INSRREQUISITN_QR(7) = Trim(txt_catgory.Text)
INSRREQUISITN_QR(8) = "8888"
INSRREQUISITN_QR(9) = UCase(Trim(cmb_refdoc.Text))
'INSRREQUISITN_QR(9) = Trim(txt_pridocser.Text)

'Added by Milind on 27-03-2006 for inserting date for the samples collected at outside
    If Check1.Value = 1 Then
        Dim Dt As String
        Dt = "18/02/1980"
        INSRREQUISITN_QR(15) = Format(CDate(Dt), "dd/mm/yyyy")
    Else
        INSRREQUISITN_QR(15) = Null
    End If
'upto here

    If SELRREQUISITNFLAG = False Then
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
        SELRREQUISITNFLAG = True
    Else
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        SELRREQUISITN_RS.MoreResults
        SELRREQUISITN_RS.Requery
    End If
    If SELRREQUISITN_RS.RowCount > 0 Then
        INSRREQUISITN_QR(10) = "YES"
    Else
        INSRREQUISITN_QR(10) = "NO"
    End If
If opt_patient.Value = True Then
    NO_CHARGE
    ''' Added by Sandhya dated 26.6.2009
    ''' for NC category & charged patient
    If txt_catgory.Text = "NC" And No_Chargeflag = True Then
        INSRREQUISITN_QR(11) = "NOCHARGE"
        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
        If Trim(txt_modcode) = "HT" Then
            INSRREQUISITN_QR(13) = "P"
        Else
            INSRREQUISITN_QR(13) = Null '"P"
        End If
        'No_Chargeflag = False
    Else
        
            INSRREQUISITN_QR(11) = Null
            INSRREQUISITN_QR(12) = Null
            INSRREQUISITN_QR(13) = Null
       
    End If
End If
If opt_staff.Value = True Then
        INSRREQUISITN_QR(11) = "NOCHARGE"
        INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
        INSRREQUISITN_QR(13) = "P"
End If
INSRREQUISITN_QR(14) = 0

INSRREQUISITN_QR.Execute
    '*** INSERTING INT0 RREQUISITION
    '*** INSERTING INTO DMULPROCE TABLE
    If CMB_TEXT.ListCount <> 0 Then
        For I = 0 To CMB_TEXT.ListCount - 1
            INSDMULPROCE_QR(0) = req_chglen(txt_bcreqno.Text)
            INSDMULPROCE_QR(1) = CMB_TEXT.List(I)
            INSDMULPROCE_QR(2) = GET_REVDATE(DATE11)
            INSDMULPROCE_QR(3) = Msk_caseno.FormattedText
            INSDMULPROCE_QR.Execute
        Next I
    End If
    '*** INSERTING INTO DMULPROCE TABLE
    
    '**** INSERTING INTO DREQUISITION

INSDREQUISITN_QR(0) = req_chglen(Trim(txt_bcreqno.Text))
INSDREQUISITN_QR(1) = " "
INSDREQUISITN_QR(2) = Trim(Msk_caseno.FormattedText)
INSDREQUISITN_QR(3) = Null

INSDREQUISITN_QR(4) = Null
dur_ill = txt_year & "/" & Trim(CMB_DURILLMONTH.Text)
INSDREQUISITN_QR(5) = dur_ill
INSDREQUISITN_QR(6) = bodyf
INSDREQUISITN_QR(7) = HLA_AGE 'Null
INSDREQUISITN_QR(8) = HLA_SEX 'Null
INSDREQUISITN_QR(9) = Null
INSDREQUISITN_QR(10) = Null
INSDREQUISITN_QR(11) = Null
INSDREQUISITN_QR(12) = Null
INSDREQUISITN_QR(13) = Null
INSDREQUISITN_QR(14) = Trim(cmb_refdoc.Text)
INSDREQUISITN_QR(15) = Trim(cmb_reftodoc.Text)
INSDREQUISITN_QR(16) = Null
INSDREQUISITN_QR(16) = HTOther
If Trim(txt_status.Text) = "IP" Then
    INSDREQUISITN_QR(17) = Trim(txt_ward.Text)
    INSDREQUISITN_QR(18) = Trim(txt_bedno.Text)
Else
    INSDREQUISITN_QR(17) = Null
    INSDREQUISITN_QR(18) = Null
End If
INSDREQUISITN_QR(19) = Null
INSDREQUISITN_QR(19) = DonorRelation
If Len(Trim(meterial)) <> 0 Then
    INSDREQUISITN_QR(20) = Null
    INSDREQUISITN_QR(20) = meterial
Else
    INSDREQUISITN_QR(20) = Null
    INSDREQUISITN_QR(20) = ""
End If
INSDREQUISITN_QR(21) = Null
INSDREQUISITN_QR(21) = STRDonor
INSDREQUISITN_QR.Execute
'**** INSERTING INTO DREQUISITION

 '*** UPDATING THE PROVISIONAL DIAGNOSIS
 
If opt_patient.Value = True Then
If SELTA1012FLAG = False Then
    selta1012_qr(0) = Msk_caseno.FormattedText
    Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
    SELTA1012FLAG = True
Else
     selta1012_qr(0) = Msk_caseno.FormattedText
     selta1012_rs.MoreResults
     selta1012_rs.Requery
     End If
If selta1012_rs.RowCount > 0 Then
    UPDTA1012_QR(0) = Trim(txt_indctinvstg.Text)
    UPDTA1012_QR(1) = Msk_caseno.FormattedText
    UPDTA1012_QR.Execute
Else
    insta1012_qr(0) = Msk_caseno.FormattedText
    insta1012_qr(1) = Trim(txt_indctinvstg.Text)
    insta1012_qr.Execute
End If
End If
'*** UPDATING THE PROVISIONAL DIAGNOSIS
'*** INSERTING INTO DMULPROCE FOR COLLECTION AND DEPT COLLECTION SCREEN
''    INSDMULPROCE_QR(0) = req_chglen(txt_spreqno.Text)
''    INSDMULPROCE_QR(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
''    INSDMULPROCE_QR(2) = Msk_caseno.FormattedText
''    INSDMULPROCE_QR(3) = cmb_procedurecode.Text
''    INSDMULPROCE_QR.Execute

'***INSERTING INTO DMULPROCE FOR COLLECTION AND DEPT COLLECTION SCREEN
'***INSERTING INTO DREQSTAFFD TABLE
If opt_staff.Value = True Then
    INSDREQSTAFFD_QR(0) = req_chglen(txt_bcreqno.Text)
    INSDREQSTAFFD_QR(1) = Msk_caseno.Text
    INSDREQSTAFFD_QR(2) = Trim(txt_name.Text)
    INSDREQSTAFFD_QR(3) = Trim(txt_Age)
    INSDREQSTAFFD_QR(4) = Trim(txt_sex)
    If Len(Trim(cmb_dependents.Text)) = 0 Then
        INSDREQSTAFFD_QR(5) = "SELF"
    Else
        I = InStr(cmb_dependents, "$")
        INSDREQSTAFFD_QR(5) = Mid(cmb_dependents, I + 1, 10)
    End If
    INSDREQSTAFFD_QR.Execute
End If
'***INSERTING INTO DREQSTAFFD TABLE
    If SELCLINHISTFLAG = False Then
        SELCLINHIST_QR(0) = Msk_caseno.FormattedText
        Set SELCLINHIST_RS = SELCLINHIST_QR.OpenResultset(1, 3)
        SELCLINHISTFLAG = True
    Else
        SELCLINHIST_QR(0) = Msk_caseno.FormattedText
        SELCLINHIST_RS.MoreResults
        SELCLINHIST_RS.Requery
    End If
'    If SELCLINHIST_RS.RowCount > 0 Then
'        updclinhist_qr(0) = Trim(txt_clinhist.Text)
'        updclinhist_qr(1) = Trim(Msk_caseno.Text)
'        updclinhist_qr.Execute
'    Else
        insclinhist_qr(0) = Trim(Msk_caseno.FormattedText)
        insclinhist_qr(1) = GET_REVDATE(Trim(DATE11))
        insclinhist_qr(2) = Trim(time1)
        If Trim(txt_clinhist.Text) = "NYD" Then
            insclinhist_qr(3) = "NYD"
        Else
            insclinhist_qr(3) = Trim(txt_clinhist.Text)
        End If
        insclinhist_qr.Execute
    'End If

Data1.UpdateRecord
Data1.Recordset.Bookmark = Data1.Recordset.LastModified
Exit Sub
MSG:
MsgBox Err.Description
End Sub


Public Sub HTVIEW()
'If SELDTMNO_FLAG = False Then
'    SELDTMNO_QR(0) = Trim(msk_caseno.FormattedText)
'    Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
'    SELDTMNO_FLAG = True
'Else
'   SELDTMNO_QR(0) = Trim(msk_caseno.FormattedText)
'   SELDTMNO_RS.MoreResults
'   SELDTMNO_RS.Requery
'End If
'If SELDTMNO_RS.RowCount > 0 Then
'    Text1.Visible = True
'    Label1.Visible = True
'    Text1.Text = Trim(SELDTMNO_RS("DTMNO"))
'    'bloodProducts.cmb_bg1.Text = Trim(SELDTMNO_RS("FLD1"))
'End If
If SELDMULPROCERS("pro_coDe") = "H102" Then
    hlatyping.chk_102.Value = 1
End If
If SELDMULPROCERS("pro_coDe") = "H103" Then
    hlatyping.chk_103.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "H104" Then
    hlatyping.chk_104.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "H101" Then
    hlatyping.CHK_101.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "H201" Then
    hlatyping.CHK_201.Value = 1
End If
If seldrequisitn_FLAG = False Then
    seldrequisitn_qr(0) = req_chglen(cmb_reqno)
    Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
    seldrequisitn_FLAG = True
Else
    seldrequisitn_qr(0) = req_chglen(cmb_reqno)
    seldrequisitn_rs.MoreResults
    seldrequisitn_rs.Requery
End If
If Not seldrequisitn_rs.EOF Then
    hlatyping.TXT_OTHERS.Text = "" & seldrequisitn_rs("OPPROCEDUR")
    hlatyping.txt_donorname.Text = "" & seldrequisitn_rs("NMETERIALS")
    hlatyping.cmb_relation.Text = "" & seldrequisitn_rs("PREPATHNO")
    'If Len(seldrequisitn_rs("labNO")) > 0 Then
    If Trim(seldrequisitn_rs("labNO")) = "yes" Then
        hlatyping.chk_donor.Value = 1
    Else
        hlatyping.chk_donor.Value = 0
    End If
    hlatyping.txt_Age.Text = "" & Trim(seldrequisitn_rs("LATERALITM"))
    hlatyping.cmb_sex.Text = "" & Trim(seldrequisitn_rs("OUTSSPECIM"))
End If
End Sub

Public Sub SSVIEW()
If SELDMULPROCERS("pro_coDe") = "H001" Then
    dtmlab.chk_BG.Value = 1
End If
If SELDMULPROCERS("pro_coDe") = "H012" Then
    dtmlab.chk_HBSAG.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "H014" Then
    dtmlab.chk_HCV.Value = 1
End If
If SELDMULPROCERS("pro_coDe") = "H013" Then
    dtmlab.chk_HIV.Value = 1
End If
If SELDMULPROCERS("pro_coDe") = "H006" Then
    dtmlab.CHK_DAT.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "H007" Then
    dtmlab.CHK_IAT.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "?" Then
    dtmlab.CHK_AS.Value = 1
End If
If SELDMULPROCERS("PRO_CODE") = "H015" Then
    dtmlab.chk_h015.Value = 1
End If
End Sub

Public Sub BPVIEW()
If SELDREQUISITNFLAG = False Then
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                    SELDREQUISITNFLAG = True
                Else
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    seldrequisitn_rs.MoreResults
                    seldrequisitn_rs.Requery
                End If
                If seldrequisitn_rs.RowCount > 0 Then
                    bloodProducts.show
                    BPFLAG = True
                    bloodProducts.Msk_caseno.Text = Msk_caseno.FormattedText
                    bloodProducts.txt_name.Text = txt_name
                    'bloodProducts.TXT_DTMNO.Text = seldrequisitn_rs("PATHNO")
'                    txt_sppathno.Visible = True
'                    txt_sppathno.Text = seldrequisitn_rs("PATHNO")
                    If SELDTMNO_FLAG = False Then
                        SELDTMNO_QR(0) = Trim(Msk_caseno.FormattedText)
                        Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                        SELDTMNO_FLAG = True
                    Else
                       SELDTMNO_QR(0) = Trim(Msk_caseno.FormattedText)
                       SELDTMNO_RS.MoreResults
                       SELDTMNO_RS.Requery
                    End If
                    If SELDTMNO_RS.RowCount > 0 Then
                        bloodProducts.txt_dtmno.Text = Trim(SELDTMNO_RS("DTMNO"))
                        bloodProducts.cmb_bg1.Text = Trim(SELDTMNO_RS("FLD1"))
                    End If
                    If SELBLOODPRODUFLAG = False Then
                        SELBLOODPRODUQR(0) = req_chglen(cmb_reqno.Text)
                        Set SELBLOODPRODURS = SELBLOODPRODUQR.OpenResultset(1, 3)
                        SELBLOODPRODUFLAG = True
                    Else
                        SELBLOODPRODUQR(0) = req_chglen(cmb_reqno.Text)
                        SELBLOODPRODURS.MoreResults
                        SELBLOODPRODURS.Requery
                    End If
                    If SELBLOODPRODURS.RowCount > 0 Then
                        bloodProducts.CMB_RT.Text = "" & SELBLOODPRODURS("RFT")
                        bloodProducts.TXT_HB.Text = "" & SELBLOODPRODURS("HB")
                        bloodProducts.TXT_DT.Text = "" & Format(Trim(SELBLOODPRODURS("REQUION")), "DD/MM/YYYY")
                        bloodProducts.TXT_TM.Text = "" & SELBLOODPRODURS("TIME")
                        bloodProducts.cmb_time.Text = "" & Trim(SELBLOODPRODURS("FLD1"))
                        bloodProducts.cmb_tr.Text = "" & Trim(SELBLOODPRODURS("TYPEOFREQ"))
                        bloodProducts.CMB_PR.Text = "" & Trim(SELBLOODPRODURS("FLD2"))
                        bloodProducts.cmb_bg1.Text = "" & Trim(SELBLOODPRODURS("BG"))
                       
                        bloodProducts.TXT_PLTC.Text = "" & Trim(SELBLOODPRODURS("FLD3"))
                        If Len(SELBLOODPRODURS("WB")) > 0 Then
                            bloodProducts.CHK_WB.Value = 1
                            bloodProducts.TXT_WB.Text = "" & SELBLOODPRODURS("WB")
                        Else
                            bloodProducts.CHK_WB.Value = 0
                            bloodProducts.TXT_WB.Text = ""
                        End If
                        If Len(SELBLOODPRODURS("PC")) > 0 Then
                            bloodProducts.CHK_PC.Value = 1
                            bloodProducts.TXT_PC.Text = "" & SELBLOODPRODURS("PC")
                        Else
                            bloodProducts.CHK_PC.Value = 0
                            bloodProducts.TXT_PC.Text = ""
                        End If
                        If Len(SELBLOODPRODURS("RDP")) > 0 Then
                            bloodProducts.CHK_RDP.Value = 1
                            bloodProducts.TXT_RDP.Text = "" & SELBLOODPRODURS("RDP")
                        Else
                            bloodProducts.CHK_RDP.Value = 0
                            bloodProducts.TXT_RDP.Text = ""
                        End If
                        If Len(SELBLOODPRODURS("SDP")) > 0 Then
                            bloodProducts.CHK_SDP.Value = 1
                            bloodProducts.TXT_SDP.Text = "" & SELBLOODPRODURS("SDP")
                        Else
                            bloodProducts.CHK_SDP.Value = 0
                            bloodProducts.TXT_SDP.Text = ""
                        End If
                        If Len(SELBLOODPRODURS("FFP")) > 0 Then
                            bloodProducts.CHK_FFP.Value = 1
                            bloodProducts.TXT_FFP.Text = "" & SELBLOODPRODURS("FFP")
                        Else
                            bloodProducts.CHK_FFP.Value = 0
                            bloodProducts.TXT_FFP.Text = ""
                        End If
                        If Len(SELBLOODPRODURS("FVIIID")) > 0 Then
                            bloodProducts.CHK_FVIIID.Value = 1
                            bloodProducts.TXT_FVIIID.Text = "" & SELBLOODPRODURS("FVIIID")
                        Else
                            bloodProducts.CHK_FVIIID.Value = 0
                            bloodProducts.TXT_FVIIID.Text = ""
                        End If
                        If Len(SELBLOODPRODURS("CRYO")) > 0 Then
                            bloodProducts.CHK_CRYO.Value = 1
                            bloodProducts.TXT_CRYO.Text = "" & SELBLOODPRODURS("CRYO")
                        Else
                            bloodProducts.CHK_CRYO.Value = 0
                            bloodProducts.TXT_CRYO.Text = ""
                        End If
                        If Len(SELBLOODPRODURS("ALB")) > 0 Then
                            bloodProducts.CHK_ALB.Value = 1
                            bloodProducts.TXT_ALB.Text = "" & SELBLOODPRODURS("ALB")
                        Else
                            bloodProducts.CHK_ALB.Value = 0
                            bloodProducts.TXT_ALB.Text = ""
                        End If
                        If Len(SELBLOODPRODURS("SR1")) > 0 Then
                            bloodProducts.CHK_IRRADIATION.Value = 1

                        Else
                            bloodProducts.CHK_IRRADIATION.Value = 0

                        End If
                        If Len(SELBLOODPRODURS("SR2")) > 0 Then
                            bloodProducts.CHK_LEUCODEPLETION.Value = 1

                        Else
                            bloodProducts.CHK_LEUCODEPLETION.Value = 0

                        End If
                        bloodProducts.TXT_OTHERS.Text = "" & SELBLOODPRODURS("OTHERS")
                        bloodProducts.txt_phname.Text = "" & SELBLOODPRODURS("PHOBOTMIST")
                        bloodProducts.TXT_NOUREPLACE = "" & SELBLOODPRODURS("NOOFUNREP")
                End If
             End If
End Sub

Public Sub dtmview()
If SELDREQUISITNFLAG = False Then
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                    SELDREQUISITNFLAG = True
                Else
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    seldrequisitn_rs.MoreResults
                    seldrequisitn_rs.Requery
                End If
                If seldrequisitn_rs.RowCount > 0 Then
                    dtmlab.Msk_caseno.Text = Msk_caseno.FormattedText
                    dtmlab.txt_name.Text = txt_name
                    TEXT1.Visible = True
                    Label1.Visible = True
                    TEXT1.Enabled = False
                    TEXT1.Text = seldrequisitn_rs("PATHNO")
                End If
                If seldrequisitn_FLAG = False Then
                    seldrequisitn_qr(0) = req_chglen(cmb_reqno)
                    Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                    seldrequisitn_FLAG = True
                Else
                    seldrequisitn_qr(0) = req_chglen(cmb_reqno)
                    seldrequisitn_rs.MoreResults
                    seldrequisitn_rs.Requery
                End If
                If Not seldrequisitn_rs.EOF Then
                    dtmlab.TXT_OTHERS.Text = "" & seldrequisitn_rs("OPPROCEDUR")
                End If
                dtmlab.show
                If SELDMULPROCEFLAG = False Then
                    SELDMULPROCEQR(0) = req_chglen(cmb_reqno.Text)
                    Set SELDMULPROCERS = SELDMULPROCEQR.OpenResultset(1, 3)
                    SELDMULPROCEFLAG = True
               Else
                    SELDMULPROCEQR(0) = req_chglen(cmb_reqno.Text)
                    SELDMULPROCERS.MoreResults
                    SELDMULPROCERS.Requery
                End If
                If SELDMULPROCERS.RowCount > 0 Then
                    While Not SELDMULPROCERS.EOF
                        SSVIEW
                        SELDMULPROCERS.MoveNext
                    Wend
                    DTMFLAG = True
                End If
End Sub

Public Sub HLAVIEW()

                If SELDREQUISITNFLAG = False Then
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                    SELDREQUISITNFLAG = True
                Else
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    seldrequisitn_rs.MoreResults
                    seldrequisitn_rs.Requery
                End If
                If seldrequisitn_rs.RowCount > 0 Then
                    hlatyping.Msk_caseno.Text = Msk_caseno.FormattedText
                    hlatyping.txt_name.Text = Trim(txt_name)
                    cmb_refdoc.Text = Trim(seldrequisitn_rs("refdoc"))
                    cmb_reftodoc.Text = Trim(seldrequisitn_rs("reftodoc"))
                End If
                hlatyping.Frame3.Enabled = False
                hlatyping.show
                If SELDMULPROCEFLAG = False Then
                    SELDMULPROCEQR(0) = req_chglen(cmb_reqno.Text)
                    Set SELDMULPROCERS = SELDMULPROCEQR.OpenResultset(1, 3)
                    SELDMULPROCEFLAG = True
               Else
                    SELDMULPROCEQR(0) = req_chglen(cmb_reqno.Text)
                    SELDMULPROCERS.MoreResults
                    SELDMULPROCERS.Requery
                End If
                If SELDMULPROCERS.RowCount > 0 Then
                    While Not SELDMULPROCERS.EOF
                        HTVIEW
                        SELDMULPROCERS.MoveNext
                    Wend
                    HLAFLAG = True
                End If
End Sub

Public Sub BODYFVIEW()
If SELDREQUISITNFLAG = False Then
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                    SELDREQUISITNFLAG = True
                Else
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    seldrequisitn_rs.MoreResults
                    seldrequisitn_rs.Requery
                End If
                If seldrequisitn_rs.RowCount > 0 Then
                    HEMBF.Msk_caseno.Text = Msk_caseno.FormattedText
                    HEMBF.txt_name.Text = txt_name
                End If
                HEMBF.show
                If SELDMULPROCEFLAG = False Then
                    SELDMULPROCEQR(0) = req_chglen(cmb_reqno.Text)
                    Set SELDMULPROCERS = SELDMULPROCEQR.OpenResultset(1, 3)
                    SELDMULPROCEFLAG = True
               Else
                    SELDMULPROCEQR(0) = req_chglen(cmb_reqno.Text)
                    SELDMULPROCERS.MoreResults
                    SELDMULPROCERS.Requery
                End If
                If SELDMULPROCERS.RowCount > 0 Then
                        While Not SELDMULPROCERS.EOF
                            BFVIEW
                            SELDMULPROCERS.MoveNext
                        Wend
                        BFFLAG = True
                End If
                    If SELDBMAINDFLAG = False Then
                        SELDBMAINDQR(0) = req_chglen(cmb_reqno.Text)
                        Set SELDBMAINDRS = SELDBMAINDQR.OpenResultset(1, 3)
                        SELDBMAINDFLAG = True
                    Else
                        SELDBMAINDQR(0) = req_chglen(cmb_reqno.Text)
                        SELDBMAINDRS.MoreResults
                        SELDBMAINDRS.Requery
                    End If
                    ' SUNDEEP CHECK
                    If SELDBMAINDRS.RowCount > 0 Then
                        If SELDBMAINDRS("NEWCASE") <> "" Then
                            HEMBF.CHK_CELLCOUNT.Value = 1
                        End If
                        If SELDBMAINDRS("POSTTRANS") <> "" Then
                            HEMBF.chk_cytomorphology.Value = 1
                        End If
                        If SELDBMAINDRS("DAY10MARR") <> "" Then
                            HEMBF.chk_cytochemistry.Value = 1
                        End If
                        If SELDBMAINDRS("PARTREATED") <> "" Then
                            HEMBF.CHK_BFC.Value = 1
                        End If
                        If SELDBMAINDRS("TSTEROIDS") <> "" Then
                            HEMBF.chk_sm.Value = 1
                        End If
                        
                    End If
End Sub

Public Sub BONEMVIEW()
Dim I
If PREPATH_FLAG = False Then
    PREPATH_QR(0) = Msk_caseno.FormattedText
    Set PREPATH_RS = PREPATH_QR.OpenResultset(1, 3)
    PREPATH_FLAG = True
Else
    PREPATH_QR(0) = Msk_caseno.FormattedText
    PREPATH_RS.MoreResults
    PREPATH_RS.Requery
End If
If Not PREPATH_RS.EOF Then
    Do While Not PREPATH_RS.EOF
        HEMBMA.cmb_bmaspno.AddItem req_chglen(PREPATH_RS("REQNNO")) & " " & "$" & PREPATH_RS("PATHNO")
        PREPATH_RS.MoveNext
    Loop
End If
If SELDREQUISITNFLAG = False Then
    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
    Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
    SELDREQUISITNFLAG = True
Else
    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
    seldrequisitn_rs.MoreResults
    seldrequisitn_rs.Requery
End If
If seldrequisitn_rs.RowCount > 0 Then
    HEMBMA.Msk_caseno.Text = Msk_caseno.FormattedText
    HEMBMA.txt_name.Text = txt_name
    I = InStr(1, Trim(seldrequisitn_rs("NMETERIALS")), "+")
    HEMBMA.cmb_nmaterial.Text = "" & Mid(Trim(seldrequisitn_rs("NMETERIALS")), 1, I - 1)
    HEMBMA.TXT_NOTHERS.Text = "" & Mid(Trim(seldrequisitn_rs("NMETERIALS")), I + 1)
End If
HEMBMA.show
If SELDMULPROCEFLAG = False Then
    SELDMULPROCEQR(0) = req_chglen(cmb_reqno.Text)
    Set SELDMULPROCERS = SELDMULPROCEQR.OpenResultset(1, 3)
    SELDMULPROCEFLAG = True
Else
    SELDMULPROCEQR(0) = req_chglen(cmb_reqno.Text)
    SELDMULPROCERS.MoreResults
    SELDMULPROCERS.Requery
End If
If SELDMULPROCERS.RowCount > 0 Then
    While Not SELDMULPROCERS.EOF
        BMVIEW
        SELDMULPROCERS.MoveNext
    Wend
    BMFLAG = True
End If
If SELDBMAINDFLAG = False Then
    SELDBMAINDQR(0) = req_chglen(cmb_reqno.Text)
    Set SELDBMAINDRS = SELDBMAINDQR.OpenResultset(1, 3)
    SELDBMAINDFLAG = True
Else
    SELDBMAINDQR(0) = req_chglen(cmb_reqno.Text)
    SELDBMAINDRS.MoreResults
    SELDBMAINDRS.Requery
End If
If SELDBMAINDRS.RowCount > 0 Then
    If SELDBMAINDRS("NEWCASE") <> "" Then
        HEMBMA.chk_nc.Value = 1
    End If
    If SELDBMAINDRS("POSTTRANS") <> "" Then
        HEMBMA.chk_pt.Value = 1
    End If
    If SELDBMAINDRS("DAY10MARR") <> "" Then
        HEMBMA.chk_dm.Value = 1
    End If
    If SELDBMAINDRS("PARTREATED") <> "" Then
        HEMBMA.chk_partiallytreated.Value = 1
    End If
    If SELDBMAINDRS("TSTEROIDS") <> "" Then
        HEMBMA.chk_ts.Value = 1
    End If
    If SELDBMAINDRS("RELAPSE") <> "" Then
        HEMBMA.chk_relapse.Value = 1
    End If
    If SELDBMAINDRS("pinduction") <> "" Then
        HEMBMA.chk_Posttreatment.Value = 1
    End If
    If SELDBMAINDRS("other") <> "" Then
        HEMBMA.CHK_OTHERS.Value = 1
        HEMBMA.TXT_OTHERS.Text = "" & Trim(SELDBMAINDRS("other"))
    End If
    If SELDBMAINDRS("POSTBMT") <> "" Then
        HEMBMA.chk_pb.Value = 1
    End If
    If SELDBMAINDRS("LYM") <> "" Then
        HEMBMA.chk_lym.Value = 1
    End If
    If SELDBMAINDRS("SPL") <> "" Then
        HEMBMA.chk_spl.Value = 1
    End If
    If SELDBMAINDRS("HEP") <> "" Then
        HEMBMA.chk_hep.Value = 1
    End If
End If
End Sub

Public Sub HEMVIEW()
If SELDHEMCYTO_FLAG = False Then
    SELDHEMCYTO_QR1(0) = req_chglen(Trim(cmb_reqno.Text))
    Set SELDHEMCYTO_RS = SELDHEMCYTO_QR1.OpenResultset(1, 3)
    SELDHEMCYTO_FLAG = True
Else
    SELDHEMCYTO_QR1(0) = req_chglen(Trim(cmb_reqno.Text))
    SELDHEMCYTO_RS.MoreResults
    SELDHEMCYTO_RS.Requery
End If
If Not SELDHEMCYTO_RS.EOF Then
    If Trim(SELDHEMCYTO_RS("MPO")) <> "" Then
        hemcbc.chk_mpo.Value = 1
    Else
        hemcbc.chk_mpo.Value = 0
    End If
   '16-04
   
    If Trim(SELDHEMCYTO_RS("hem")) <> "" Then
        hemcbc.CHK_HAEMOGRAM.Value = 1
    Else
        hemcbc.CHK_HAEMOGRAM.Value = 0
    End If
    If Trim(SELDHEMCYTO_RS("SMAR")) <> "" Then
       '''
         If SELDMULPROCEFLAG = False Then
                    SELDMULPROCEQR(0) = req_chglen(cmb_reqno.Text)
                    Set SELDMULPROCERS = SELDMULPROCEQR.OpenResultset(1, 3)
                    SELDMULPROCEFLAG = True
               Else
                    SELDMULPROCEQR(0) = req_chglen(cmb_reqno.Text)
                    SELDMULPROCERS.MoreResults
                  SELDMULPROCERS.Requery
         End If
       If SELDMULPROCERS.RowCount > 0 Then
            SELDMULPROCERS.MoveFirst
         Do While SELDMULPROCERS.EOF = False
             If SELDMULPROCERS("pro_code") = "U753" Then
                hemcbc.CHK_SMC.Value = 1
             Else
                 hemcbc.CHK_SMC.Value = 0
             End If
             If SELDMULPROCERS("pro_code") = "U754" Then
                hemcbc.chk_sm.Value = 1
             Else
                 hemcbc.chk_sm.Value = 0
             End If
           SELDMULPROCERS.MoveNext
         Loop
       End If

       ''''
        'hemcbc.CHK_HAEMOGRAM.Value = 1
   ' Else
        'hemcbc.CHK_HAEMOGRAM.Value = 0
    End If
    If Trim(SELDHEMCYTO_RS("CBC")) <> "" Then
        hemcbc.chk_cbc.Value = 1
    Else
        hemcbc.chk_cbc.Value = 0
    End If
    
    '16-04
    If Trim(SELDHEMCYTO_RS("SBB")) <> "" Then
        hemcbc.CHK_SBB.Value = 1
    Else
        hemcbc.CHK_SBB.Value = 0
    End If
    If Trim(SELDHEMCYTO_RS("PAS")) <> "" Then
        hemcbc.CHK_PAS.Value = 1
    Else
        hemcbc.CHK_PAS.Value = 0
    End If
    If Trim(SELDHEMCYTO_RS("LAP")) <> "" Then
        hemcbc.CHK_LAP.Value = 1
    Else
        hemcbc.CHK_LAP.Value = 0
    End If
    If Trim(SELDHEMCYTO_RS("NSE")) <> "" Then
        hemcbc.CHK_NSE.Value = 1
    Else
        hemcbc.CHK_NSE.Value = 0
    End If
    If Trim(SELDHEMCYTO_RS("TRAP")) <> "" Then
        hemcbc.CHK_TRAP1.Value = 1
    Else
        hemcbc.CHK_TRAP1.Value = 0
    End If
    If Trim(SELDHEMCYTO_RS("fld1")) <> "" Then
        hemcbc.CHK_TT.Value = 1
    Else
        hemcbc.CHK_TT.Value = 0
    End If
    If Trim(SELDHEMCYTO_RS("fld2")) <> "" Then
        hemcbc.CHK_OTHERS.Value = 1
        hemcbc.TEXT1.Text = Trim(SELDHEMCYTO_RS("fld3"))
    Else
        hemcbc.CHK_OTHERS.Value = 0
        hemcbc.TEXT1.Text = Trim(SELDHEMCYTO_RS("fld3"))
    End If
End If
                If SELDREQUISITNFLAG = False Then
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                    SELDREQUISITNFLAG = True
                Else
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    seldrequisitn_rs.MoreResults
                    seldrequisitn_rs.Requery
                End If
                If seldrequisitn_rs.RowCount > 0 Then
                    hemcbc.Msk_caseno.Text = Msk_caseno.FormattedText
                    hemcbc.txt_name.Text = txt_name
                End If
                hemcbc.show
                If PREPATH_FLAG = False Then
                    PREPATH_QR(0) = Msk_caseno.FormattedText
                    Set PREPATH_RS = PREPATH_QR.OpenResultset(1, 3)
                    PREPATH_FLAG = True
                Else
                    PREPATH_QR(0) = Msk_caseno.FormattedText
                    PREPATH_RS.MoreResults
                    PREPATH_RS.Requery
                End If
                If Not PREPATH_RS.EOF Then
                    Do While Not PREPATH_RS.EOF
                        hemcbc.cmb_preinvest.AddItem req_chglen(PREPATH_RS("REQNNO")) & " " & "$" & PREPATH_RS("PATHNO")
                        PREPATH_RS.MoveNext
                    Loop
                End If
                If SELDMULPROCEFLAG = False Then
                    SELDMULPROCEQR(0) = req_chglen(cmb_reqno.Text)
                    Set SELDMULPROCERS = SELDMULPROCEQR.OpenResultset(1, 3)
                    SELDMULPROCEFLAG = True
               Else
                    SELDMULPROCEQR(0) = req_chglen(cmb_reqno.Text)
                    SELDMULPROCERS.MoreResults
                    SELDMULPROCERS.Requery
                End If
                If SELDMULPROCERS.RowCount > 0 Then
                        While Not SELDMULPROCERS.EOF
                            HLVIEW
                            SELDMULPROCERS.MoveNext
                        Wend
                        HLFLAG = True
                End If
End Sub

Public Sub CARDVIEW()

 frm_hematopathology.Visible = True

                
                   lbl_siteed.Visible = False
                txt_edsite.Visible = False
                lbl_biopsycytology.Visible = False
                'cmb_biopsycytology.Visible = False
                lbl_requestremarks.Visible = False
                txt_REQUESTDETAILS.Visible = False
                lbl_edprerequisition.Visible = False
                cmb_edprereqno.Visible = False
                lbl_edprocedureheading.Visible = False
                cmb_edprocedureheading.Visible = False
                list_edts.Visible = False
                cmb_hpprocedure.Enabled = False
                cmb_hpprocedurecode.Enabled = False
                txt_hpreqno.Visible = False
                lbl_hpreqno1.Visible = False
                
                If Trim(txt_modcode.Text) = "EK" Then
                    cmd_edindication.Enabled = False
                Else
                    cmd_edindication.Enabled = True
                End If
                cmd_edindication.Enabled = False
                If SELRREQUISITNFLAG2 = False Then
                     SELRREQUISITN_QR2(0) = req_chglen(Trim(cmb_reqno.Text))
                     Set SELRREQUISITN_RS2 = SELRREQUISITN_QR2.OpenResultset(1, 3)
                     SELRREQUISITNFLAG2 = True
                Else
                     SELRREQUISITN_QR2(0) = req_chglen(Trim(cmb_reqno.Text))
                     SELRREQUISITN_RS2.MoreResults
                     SELRREQUISITN_RS2.Requery
                End If
                If SELRREQUISITN_RS2.RowCount > 0 Then
                    cmb_hpprocedure.clear
                    cmb_hpprocedurecode.clear
                
                        cmb_hpprocedurecode.AddItem SELRREQUISITN_RS2("SER_CODE")
                        cmb_hpprocedurecode.Text = SELRREQUISITN_RS2("SER_CODE")
                        If SELMA101FLAG = False Then
                            SELMA101_QR(0) = Trim(SELRREQUISITN_RS2("SER_CODE"))
                            Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                            SELMA101FLAG = True
                        Else
                            SELMA101_QR(0) = Trim(SELRREQUISITN_RS2("SER_CODE"))
                            SELMA101_RS.MoreResults
                            SELMA101_RS.Requery
                        End If
                        If SELMA101_RS.RowCount > 0 Then
                            While Not SELMA101_RS.EOF
                                cmb_hpprocedure.AddItem Trim(SELMA101_RS("SER_DET"))
                                cmb_hpprocedure.Text = Trim(SELMA101_RS("SER_DET"))
                                SELMA101_RS.MoveNext
                            Wend
                         End If
                     
                End If
                
                If SELINDICATIONFLAG = False Then
                    SELINDICATIONQR(0) = req_chglen(Trim(cmb_reqno.Text))
                    Set SELINDICATIONRS = SELINDICATIONQR.OpenResultset(1, 3)
                    SELINDICATIONFLAG = False
                Else
                    SELINDICATIONQR(0) = req_chglen(Trim(cmb_reqno.Text))
                    SELINDICATIONRS.MoreResults
                    SELINDICATIONRS.Requery
                End If
                If SELINDICATIONRS.RowCount > 0 Then
                    ECVIEW
                
                End If
End Sub

Public Sub ENDOSVIEW()
frm_hematopathology.Visible = True
                EDENABL ("DIS")
                EDCLEAR
                txt_hpreqno.Visible = False
                lbl_hpreqno1.Visible = False
                TXT_MBPATHNO.Visible = False
                'LBL_MBMICRONO.Visible = False
                'txt_pathno.Visible = True
                'lbl_pathno.Visible = True
                cmd_edindication.Visible = False
                cmb_edprereqno.Visible = False
                lbl_edprerequisition.Visible = False
                lbl_biopsycytology.Visible = True
                list_edts.Visible = True
                cmb_edprocedureheading.Visible = True
                lbl_edprocedureheading.Visible = True
                lbl_requestremarks.Visible = True
                txt_REQUESTDETAILS.Visible = True
                cmb_edprereqno.Visible = True
                lbl_edprerequisition.Visible = True
                
                
                I1 = 0
                If SELRREQUISITNFLAG2 = False Then
                     SELRREQUISITN_QR2(0) = req_chglen(Trim(cmb_reqno.Text))
                     Set SELRREQUISITN_RS2 = SELRREQUISITN_QR2.OpenResultset(1, 3)
                     SELRREQUISITNFLAG2 = True
                Else
                     SELRREQUISITN_QR2(0) = req_chglen(Trim(cmb_reqno.Text))
                     SELRREQUISITN_RS2.MoreResults
                     SELRREQUISITN_RS2.Requery
                End If
                If SELRREQUISITN_RS2.RowCount > 0 Then
                    cmb_hpprocedure.clear
                    cmb_hpprocedurecode.clear
                
                        cmb_hpprocedurecode.AddItem SELRREQUISITN_RS2("SER_CODE")
                        cmb_hpprocedurecode.Text = SELRREQUISITN_RS2("SER_CODE")
                        If SELMA101FLAG = False Then
                            SELMA101_QR(0) = Trim(SELRREQUISITN_RS2("SER_CODE"))
                            Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                            SELMA101FLAG = True
                        Else
                            SELMA101_QR(0) = Trim(SELRREQUISITN_RS2("SER_CODE"))
                            SELMA101_RS.MoreResults
                            SELMA101_RS.Requery
                        End If
                        If SELMA101_RS.RowCount > 0 Then
                            While Not SELMA101_RS.EOF
                                cmb_hpprocedure.AddItem Trim(SELMA101_RS("SER_DET"))
                                cmb_hpprocedure.Text = Trim(SELMA101_RS("SER_DET"))
                                SELMA101_RS.MoveNext
                            Wend
                         End If
                        
                        
                        
                End If
                
                
                
                If SELDREQUISITNFLAG = False Then
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                    SELDREQUISITNFLAG = True
                Else
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    seldrequisitn_rs.MoreResults
                    seldrequisitn_rs.Requery
                End If
                    If seldrequisitn_rs.RowCount > 0 Then
                       cmb_refdoc.Text = Trim(seldrequisitn_rs("REFDOC"))
                       cmb_reftodoc.Text = Trim(seldrequisitn_rs("reftodoc"))
                       lbl_pathno.Visible = False
                       txt_pathno.Visible = False
                       'txt_pathno.Text = Trim(seldrequisitn_rs("pathno"))
                       I = InStr(seldrequisitn_rs("durillness"), "/")
                       If I <> 0 Then
                        txt_year.Text = Mid(Trim(seldrequisitn_rs("durillness")), 1, I - 1)
                        CMB_DURILLMONTH.Text = Mid(Trim(seldrequisitn_rs("durillness")), I + 1, 2)
                       End If
                       If seldrequisitn_rs.EOF = False Then
                            txt_edsite.Text = " " & Trim(seldrequisitn_rs("NMETERIALS"))
                       End If
                       txt_REQUESTDETAILS.Text = Trim(seldrequisitn_rs("PREPATHNO"))
                       cmb_edprocedureheading.Text = Trim(seldrequisitn_rs("METASTASIS"))
                       list_edts.AddItem Trim(seldrequisitn_rs("OUTSSPECIM"))
                       list_edts.Text = Trim(seldrequisitn_rs("OUTSSPECIM"))
                    End If
                
End Sub

Public Sub MBVIEW()
Microbiology.Visible = True
                MBENABL ("DIS")
                MBCLEAR
                txt_mbreqno.Visible = False
                lbl_mbreqno.Visible = False
                TXT_MBPATHNO.Visible = False
                LBL_MBMICRONO.Visible = False
                txt_pathno.Visible = True
                lbl_pathno.Visible = True
                I1 = 0
                If SELDMULPROCEFLAG = False Then
                    SELDMULPROCEQR(0) = req_chglen(cmb_reqno.Text)
                    Set SELDMULPROCERS = SELDMULPROCEQR.OpenResultset(1, 3)
                    SELDMULPROCEFLAG = True
               Else
                    SELDMULPROCEQR(0) = req_chglen(cmb_reqno.Text)
                    SELDMULPROCERS.MoreResults
                    SELDMULPROCERS.Requery
                End If
                If SELDMULPROCERS.RowCount > 0 Then
                    List_code.clear
                    list_proc.clear
                    While Not SELDMULPROCERS.EOF
                        List_code.AddItem SELDMULPROCERS("PRO_CODE")
                        
                        If SELMA101FLAG = False Then
                            SELMA101_QR(0) = Trim(SELDMULPROCERS("pro_CODE"))
                            Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                            SELMA101FLAG = True
                        Else
                            SELMA101_QR(0) = Trim(SELDMULPROCERS("pro_CODE"))
                            SELMA101_RS.MoreResults
                            SELMA101_RS.Requery
                        End If
                        If SELMA101_RS.RowCount > 0 Then
                            While Not SELMA101_RS.EOF
                                list_proc.AddItem SELMA101_RS("SER_DET")
                                
                                SELMA101_RS.MoveNext
                            Wend
                         End If
                         SELDMULPROCERS.MoveNext
                    Wend
                        For I = 0 To List_code.ListCount - 1
                            List_code.Selected(I) = True
                            
                        Next I
                        For I = 0 To list_proc.ListCount - 1
                            list_proc.Selected(I) = True
                        Next I
                        
                End If
                
                
                
                If SELDREQUISITNFLAG = False Then
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                    SELDREQUISITNFLAG = True
                Else
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    seldrequisitn_rs.MoreResults
                    seldrequisitn_rs.Requery
                End If
                
                    cmb_refdoc.Text = Trim(seldrequisitn_rs("REFDOC"))
                    cmb_reftodoc.Text = Trim(seldrequisitn_rs("reftodoc"))
                    lbl_pathno.Visible = True
                    txt_pathno.Visible = True
                    txt_pathno.Text = Trim(seldrequisitn_rs("pathno"))
                    'CMB_ICU.Text = "" & seldrequisitn_rs("LABNO")
                    I = InStr(seldrequisitn_rs("durillness"), "/")
                    If I <> 0 Then
                    txt_year.Text = Mid(Trim(seldrequisitn_rs("durillness")), 1, I - 1)
                    CMB_DURILLMONTH.Text = Mid(Trim(seldrequisitn_rs("durillness")), I + 1, 2)
                    End If
                    If seldrequisitn_rs.EOF = False Then
                        cmb_mboutslide.Text = " " & (seldrequisitn_rs("OUTSLIDES"))
                    Else
                        cmb_mboutslide.Text = ""
                    End If
                    If SELSPONANTIBFLAG = False Then
                        SELSPONANTIBQR(0) = req_chglen(cmb_reqno.Text)
                        Set SELSPONANTIBRS = SELSPONANTIBQR.OpenResultset(1, 3)
                        SELSPONANTIBFLAG = True
                    Else
                        SELSPONANTIBQR(0) = req_chglen(cmb_reqno.Text)
                        SELSPONANTIBRS.MoreResults
                        SELSPONANTIBRS.Requery
                    End If
                    If SELSPONANTIBRS.RowCount > 0 Then
                        If SELSPONANTIBRS.EOF = False Then
                            'txt_mbspecimen.Text = Trim(SELSPONANTIBRS("SPECIMEN").Value)
                            cmb_specimen.Text = Trim(SELSPONANTIBRS("SPECIMEN").Value)
                        Else
                            'txt_mbspecimen.Text = ""
                            cmb_specimen.Text = ""
                        End If
                        If SELSPONANTIBRS.EOF = False Then
                            txt_ongoinganti.Text = Trim(SELSPONANTIBRS("ONGOINGANT").Value)
                        Else
                            txt_ongoinganti.Text = ""
                        End If
                    End If
End Sub

Public Sub BIOVIEW()
If SELDREQUISITNFLAG = False Then
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                    SELDREQUISITNFLAG = True
                Else
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    seldrequisitn_rs.MoreResults
                    seldrequisitn_rs.Requery
                End If
                If seldrequisitn_rs.RowCount > 0 Then
                    BIOCHEMISTRY1.Msk_caseno.Text = Msk_caseno.FormattedText
                    BIOCHEMISTRY1.CMB_BODYF.Text = Trim(seldrequisitn_rs("METASTASIS"))
                    BIOCHEMISTRY1.txt_name.Text = txt_name
                    BIOCHEMISTRY1.CMB_NMETERIAL.Text = Trim(seldrequisitn_rs("nmeterials"))
                End If
                BIOCHEMISTRY1.show
                If Trim(BIOCHEMISTRY1.CMB_NMETERIAL.Text) = "Blood" Then
                    BIOCHEMISTRY1.blood.Visible = True
                    BIOCHEMISTRY1.urinerandom.Visible = False
                    BIOCHEMISTRY1.urine24hour.Visible = False
                    BIOCHEMISTRY1.bodyfluid.Visible = False
                    BIOCHEMISTRY1.STONEANALYSIS.Visible = True
                    BIOCHEMISTRY1.FRMGTT.Visible = False
                ElseIf Trim(BIOCHEMISTRY1.CMB_NMETERIAL.Text) = "Stone Analysis" Then
                    BIOCHEMISTRY1.blood.Visible = False
                    BIOCHEMISTRY1.urinerandom.Visible = False
                    BIOCHEMISTRY1.urine24hour.Visible = False
                    BIOCHEMISTRY1.bodyfluid.Visible = False
                    BIOCHEMISTRY1.STONEANALYSIS.Visible = True
                    BIOCHEMISTRY1.FRMGTT.Visible = False
                ElseIf Trim(BIOCHEMISTRY1.CMB_NMETERIAL.Text) = "Body Fluid" Then
                    BIOCHEMISTRY1.blood.Visible = False
                    BIOCHEMISTRY1.urinerandom.Visible = False
                    BIOCHEMISTRY1.urine24hour.Visible = False
                    BIOCHEMISTRY1.bodyfluid.Visible = True
                    BIOCHEMISTRY1.STONEANALYSIS.Visible = False
                    BIOCHEMISTRY1.FRMGTT.Visible = False
                ElseIf Trim(BIOCHEMISTRY1.CMB_NMETERIAL.Text) = "Random Urine" Then
                    BIOCHEMISTRY1.blood.Visible = False
                    BIOCHEMISTRY1.urinerandom.Visible = True
                    BIOCHEMISTRY1.urine24hour.Visible = False
                    BIOCHEMISTRY1.bodyfluid.Visible = False
                    BIOCHEMISTRY1.STONEANALYSIS.Visible = False
                    BIOCHEMISTRY1.FRMGTT.Visible = False
                ElseIf Trim(BIOCHEMISTRY1.CMB_NMETERIAL.Text) = "24 hours Urine" Then
                    BIOCHEMISTRY1.blood.Visible = False
                    BIOCHEMISTRY1.urinerandom.Visible = False
                    BIOCHEMISTRY1.urine24hour.Visible = True
                    BIOCHEMISTRY1.bodyfluid.Visible = False
                    BIOCHEMISTRY1.STONEANALYSIS.Visible = False
                    BIOCHEMISTRY1.FRMGTT.Visible = False
                ElseIf Trim(BIOCHEMISTRY1.CMB_NMETERIAL.Text) = "GTT" Then
                    BIOCHEMISTRY1.blood.Visible = False
                    BIOCHEMISTRY1.urinerandom.Visible = False
                    BIOCHEMISTRY1.urine24hour.Visible = False
                    BIOCHEMISTRY1.bodyfluid.Visible = False
                    BIOCHEMISTRY1.STONEANALYSIS.Visible = False
                    BIOCHEMISTRY1.FRMGTT.Visible = True
                End If
                If SELDMULPROCEFLAG = False Then
                    SELDMULPROCEQR(0) = req_chglen(cmb_reqno.Text)
                    Set SELDMULPROCERS = SELDMULPROCEQR.OpenResultset(1, 3)
                    SELDMULPROCEFLAG = True
               Else
                    SELDMULPROCEQR(0) = req_chglen(cmb_reqno.Text)
                    SELDMULPROCERS.MoreResults
                    SELDMULPROCERS.Requery
                End If
                'H1
                If SELDMULPROCERS.RowCount > 0 Then
                        While Not SELDMULPROCERS.EOF
                            BCVIEW
                            SELDMULPROCERS.MoveNext
                        Wend
                        BCFLAG = True
                End If
End Sub

Public Sub EMERVIEW()
If SELDREQUISITNFLAG = False Then
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                    SELDREQUISITNFLAG = True
                Else
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    seldrequisitn_rs.MoreResults
                    seldrequisitn_rs.Requery
                End If
                If seldrequisitn_rs.RowCount > 0 Then
                  Em_lab.MaskEdBox1 = Msk_caseno.FormattedText
                    
                    Em_lab.txt_name.Text = txt_name
                End If
                Em_lab.show
                If SELDMULPROCEFLAG = False Then
                    SELDMULPROCEQR(0) = req_chglen(cmb_reqno.Text)
                    Set SELDMULPROCERS = SELDMULPROCEQR.OpenResultset(1, 3)
                    SELDMULPROCEFLAG = True
               Else
                    SELDMULPROCEQR(0) = req_chglen(cmb_reqno.Text)
                    SELDMULPROCERS.MoreResults
                    SELDMULPROCERS.Requery
                End If
                If SELDMULPROCERS.RowCount > 0 Then
                    While Not SELDMULPROCERS.EOF
                        ELVIEW
                        SELDMULPROCERS.MoveNext
                    Wend
                    ELFLAG = True
                End If
End Sub

Public Sub TUMOURVIEW()
If SELDREQUISITNFLAG = False Then
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                    SELDREQUISITNFLAG = True
                Else
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    seldrequisitn_rs.MoreResults
                    seldrequisitn_rs.Requery
                End If
                If seldrequisitn_rs.RowCount > 0 Then
                    TUMOURMARKER.txt_clininfo = txt_clinhist
                    TUMOURMARKER.txt_Dosage = seldrequisitn_rs("PREPATHNO")
                    If seldrequisitn_rs.EOF = False Then
                        TUMOURMARKER.CMB_NMETERIAL = " " & seldrequisitn_rs("NMETERIALS")
                    Else
                        TUMOURMARKER.CMB_NMETERIAL = ""
                    End If
                    TUMOURMARKER.txt_rem = seldrequisitn_rs("METASTASIS")
                End If
                TUMOURMARKER.Msk_caseno.Text = Msk_caseno.FormattedText
                TUMOURMARKER.txt_name = txt_name.Text
                TUMOURMARKER.show
                If SELDMULPROCEFLAG = False Then
                    SELDMULPROCEQR(0) = req_chglen(cmb_reqno.Text)
                    Set SELDMULPROCERS = SELDMULPROCEQR.OpenResultset(1, 3)
                    SELDMULPROCEFLAG = True
               Else
                    SELDMULPROCEQR(0) = req_chglen(cmb_reqno.Text)
                    SELDMULPROCERS.MoreResults
                    SELDMULPROCERS.Requery
                End If
                If SELDMULPROCERS.RowCount > 0 Then
                    
                    While Not SELDMULPROCERS.EOF
                        TMVIEW
                        SELDMULPROCERS.MoveNext
                    Wend
                    
                    'TUMOURMARKER.CHKCLEAR
                    TUMOURMARKER.Frame1.Enabled = False
                    TUMOURMARKER.Frame2.Enabled = False
                    TUMOURMARKER.Frame3.Enabled = False
                    tmflag = True
                End If
End Sub

Public Sub CGVIEW()
CYTOLOGY.Visible = True
                CGENABL ("DIS")
                txt_cyreqno1.Visible = False
                lbl_cyreqno.Visible = False
                
                cbo_proccode.Text = SELRREQUISITN_RS2("SER_CODE")
                
                
                If SELMA101FLAG = False Then
                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS2("SER_CODE"))
                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                    SELMA101FLAG = True
                Else
                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS2("SER_CODE"))
                    SELMA101_RS.MoreResults
                    SELMA101_RS.Requery
                End If
                If SELMA101_RS.RowCount > 0 Then
                    CMB_CGPROCEDURE.Text = SELMA101_RS("SER_DET")
                End If
                If SELRREQUISITN_RS2.RowCount > 0 Then
                    If SELRREQUISITN_RS2("RECORDNO") = 1 Then
                    CMB_CGNIP.Text = "VPP"
                    End If
                    If SELRREQUISITN_RS2("RECORDNO") = 2 Then
                    CMB_CGNIP.Text = "LR"
                    End If
                    If SELRREQUISITN_RS2("RECORDNO") = 3 Then
                    CMB_CGNIP.Text = "BAR"
                    End If
                    If SELRREQUISITN_RS2("RECORDNO") = 4 Then
                    CMB_CGNIP.Text = "DrOp"
                    End If
                End If
                If SELDREQUISITNFLAG = False Then
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                    SELDREQUISITNFLAG = True
                Else
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    seldrequisitn_rs.MoreResults
                    seldrequisitn_rs.Requery
                End If
                If seldrequisitn_rs.RowCount > 0 Then
                    If SELMA301FLAG = False Then
                        SELMA301QR(0) = seldrequisitn_rs("SITE")
                        Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
                        SELMA301FLAG = True
                    Else
                        SELMA301QR(0) = seldrequisitn_rs("SITE")
                        Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
                        SELMA301RS.MoreResults
                        SELMA301RS.Requery
                
                    End If
                    If SELMA301RS.RowCount > 0 Then
                        If Trim(SELMA301RS("SITEDESC")) = "site not specified/recorded" Then
                            cmb_cnsite.Text = ""
                        Else
                            cmb_cnsite.Text = SELMA301RS("SITEDESC")
                        End If
                    Else
                            cmb_cnsite.Text = ""
                    End If
                    cmb_refdoc.Text = Trim(seldrequisitn_rs("REFDOC"))
                    cmb_claterality.Text = Trim(seldrequisitn_rs("LATERALITS"))
                    txt_cmetastasis.Text = Trim(seldrequisitn_rs("METASTASIS"))
                    
                    'cmb_cyoutspecimen.Text = Trim(seldrequisitn_rs("OUTSSPECIM"))
                    cmb_cnoslides.Text = Trim(seldrequisitn_rs("OUTBLOCK"))
                    cmb_coutslides.Text = Trim(seldrequisitn_rs("OUTSLIDES"))
                    If seldrequisitn_rs.EOF = False Then
                        txt_cnatmeterial.Text = " " & Trim(seldrequisitn_rs("NMETERIALS"))
                    Else
                        txt_cnatmeterial.Text = " "
                    End If
                    
                    If txt_modcode.Text = "CG" Then
                        lbl_lmp.Visible = True
                        MSK_LMP.Visible = True
                        If IsNull(Trim(seldrequisitn_rs("FROZENDATE"))) Then
                        Else
                        MSK_LMP.Text = Trim(seldrequisitn_rs("FROZENDATE"))
                        End If
                    Else
                        lbl_lmp.Visible = False
                        MSK_LMP.Visible = False
                    End If
                    
                    If SELSCOPYFNDFLAG = False Then
                        SELSCOPYFNDQR(0) = Msk_caseno.FormattedText
                        Set SELSCOPYFNDRS = SELSCOPYFNDQR.OpenResultset(1, 3)
                        SELSCOPYFNDFLAG = True
                    Else
                        SELSCOPYFNDQR(0) = Msk_caseno.FormattedText
                        SELSCOPYFNDRS.MoreResults
                        SELSCOPYFNDRS.Requery
                    End If
                    
                    If SELSCOPYFNDRS.RowCount > 0 Then
                        If SELSCOPYFNDRS.EOF = False Then
                            txt_scopyfind.Text = SELSCOPYFNDRS("SURGPROCNO")
                        Else
                            txt_scopyfind.Text = ""
                        End If
                    End If
                    
                    If seldsftdcyflag = False Then
                        seldsftdcyqr(0) = req_chglen(cmb_reqno.Text)
                        Set seldsftdcyrs = seldsftdcyqr.OpenResultset(1, 3)
                        seldsftdcyflag = True
                    Else
                        seldsftdcyqr(0) = req_chglen(cmb_reqno.Text)
                        seldsftdcyrs.MoreResults
                        seldsftdcyrs.Requery
                    End If
                    If seldsftdcyrs.RowCount > 0 Then
                        
'                        If seldsftdcyrs.EOF = False Then
'                            txt_scopyfind.Text = seldsftdcyrs("spfind")
'                         Else
'                            txt_scopyfind.Text = ""
'                        End If
                        
                        If seldsftdcyrs.EOF = False Then
                            txt_trtdetail.Text = " " & seldsftdcyrs("trdet")
                         Else
                            txt_trtdetail.Text = ""
                        End If
                        If seldsftdcyrs.EOF = False Then
                            cmb_cyoutspecimen.Text = seldsftdcyrs("outspec")
                         Else
                            cmb_cyoutspecimen.Text = ""
                        End If
                    End If
                End If
End Sub

Public Sub FNACVIEW()
fnac.Visible = True
                FAENABL ("DIS")
                If Left(Msk_caseno.FormattedText, 2) = "RF" Or Left(Msk_caseno.FormattedText, 2) = "RG" Then
                        cmb_outsideslidesfn.Visible = True
                        lbl_outsideslides.Visible = True
                Else
                        cmb_outsideslidesfn.Visible = False
                        lbl_outsideslides.Visible = False
                End If
                txt_FNreqno.Visible = False
                lbl_fnreqno.Visible = False
                cmb_fnprocode.Text = SELRREQUISITN_RS2("SER_CODE")
                
                
                If SELMA101FLAG = False Then
                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS2("SER_CODE"))
                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                    SELMA101FLAG = True
                Else
                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS2("SER_CODE"))
                    SELMA101_RS.MoreResults
                    SELMA101_RS.Requery
                End If
                If SELMA101_RS.RowCount > 0 Then
                    CMB_FNPROCEDURE.Text = SELMA101_RS("SER_DET")
                End If
                If SELRREQUISITN_RS2.RowCount > 0 Then
                    If SELRREQUISITN_RS2("RECORDNO") = 1 Then
                    CMB_FANIP.Text = "VPP"
                    End If
                    If SELRREQUISITN_RS2("RECORDNO") = 2 Then
                    CMB_FANIP.Text = "LR"
                    End If
                    If SELRREQUISITN_RS2("RECORDNO") = 3 Then
                    CMB_FANIP.Text = "BAR"
                    End If
                    If SELRREQUISITN_RS2("RECORDNO") = 4 Then
                    CMB_FANIP.Text = "DrOp"
                    End If
                End If
                If SELDREQUISITNFLAG = False Then
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                    SELDREQUISITNFLAG = True
                Else
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    seldrequisitn_rs.MoreResults
                    seldrequisitn_rs.Requery
                End If
                If seldrequisitn_rs.RowCount > 0 Then
                    If SELMA301FLAG = False Then
                        SELMA301QR(0) = seldrequisitn_rs("SITE")
                        Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
                        SELMA301FLAG = True
                    Else
                        SELMA301QR(0) = seldrequisitn_rs("SITE")
                        Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
                        SELMA301RS.MoreResults
                        SELMA301RS.Requery
                
                    End If
                    If SELMA301RS.RowCount > 0 Then
                        If Trim(SELMA301RS("SITEDESC")) = "site not specified/recorded" Then
                            cmb_fnsite.Text = ""
                        Else
                            cmb_fnsite.Text = SELMA301RS("SITEDESC")
                        End If
                    Else
                            cmb_fnsite.Text = ""
                    End If
                    cmb_refdoc.Text = Trim(seldrequisitn_rs("REFDOC"))
                    cmb_reftodoc.Text = Trim(seldrequisitn_rs("reftodoc"))
                    I = InStr(seldrequisitn_rs("durillness"), "/")
                    If I <> 0 Then
                    txt_year.Text = Mid(Trim(seldrequisitn_rs("durillness")), 1, I - 1)
                    CMB_DURILLMONTH.Text = Mid(Trim(seldrequisitn_rs("durillness")), I + 1, 2)
                    End If
                    cmb_fnlaterality.Text = Trim(seldrequisitn_rs("LATERALITS"))
                    txt_fnmetastasis.Text = Trim(seldrequisitn_rs("METASTASIS"))
                    cmb_fnlaterality1.Text = Trim(seldrequisitn_rs("LATERALITM"))
                    TXT_FNSLIDEDRY.Text = Trim(seldrequisitn_rs("OUTSSPECIM"))
                    TXT_FNOSLIDEWET.Text = Trim(seldrequisitn_rs("OUTBLOCK"))
                    cmb_outsideslidesfn.Text = Trim(seldrequisitn_rs("OUTSLIDES"))
                    If seldrequisitn_rs.EOF = False Then
                        txt_fnnmeterial.Text = " " & seldrequisitn_rs("NMETERIALS").Value
                    Else
                        txt_fnnmeterial.Text = ""
                    End If
                    cmb_fnac.Text = Trim(seldrequisitn_rs("OPPROCEDUR"))
                    If Trim(seldrequisitn_rs("frozensec")) = "" Then
                    Else
'                    CMB_CELLBLOCK.AddItem Trim(seldrequisitn_rs("FROZENSEC"))
'                    CMB_CELLBLOCK.Text = Trim(seldrequisitn_rs("FROZENSEC"))
                    End If
                End If
End Sub

Public Sub SPVIEW()
lbl_pathno.Visible = True
                'txt_pathno.Visible = True
                'txt_pathno.Enabled = False
                'lbl_sppathno1.Visible = False
            
                surgicalPathology.Visible = True
                
                cmb_procedurecode.Text = SELRREQUISITN_RS2("SER_CODE")
                If SELRREQUISITN_RS2.RowCount > 0 Then
                    If SELRREQUISITN_RS2("RECORDNO") = 1 Then
                    cmb_vpp.Text = "VPP"
                    End If
                    If SELRREQUISITN_RS2("RECORDNO") = 2 Then
                    cmb_vpp.Text = "DER"
                    End If
                    If SELRREQUISITN_RS2("RECORDNO") = 3 Then
                    cmb_vpp.Text = "BAR"
                    End If
                    If SELRREQUISITN_RS2("RECORDNO") = 4 Then
                    cmb_vpp.Text = "DrOp"
                    End If
''                    If SELRREQUISITN_RS2("RECORDNO") = 5 Then
''                    cmb_vpp.Text = "CHIP"
''                    End If
                End If
                
                If SELMA101FLAG = False Then
                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS2("SER_CODE"))
                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                    SELMA101FLAG = True
                Else
                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS2("SER_CODE"))
                    SELMA101_RS.MoreResults
                    SELMA101_RS.Requery
                End If
                If SELMA101_RS.RowCount > 0 Then
                    cmb_procedure.Text = SELMA101_RS("SER_DET")
                End If
                If SELDREQUISITNFLAG = False Then
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                    SELDREQUISITNFLAG = True
                Else
                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
                    seldrequisitn_rs.MoreResults
                    seldrequisitn_rs.Requery
                End If
                If seldrequisitn_rs.RowCount > 0 Then
                    If SELMA301FLAG = False Then
                        SELMA301QR(0) = seldrequisitn_rs("SITE")
                        Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
                        SELMA301FLAG = True
                    Else
                        SELMA301QR(0) = seldrequisitn_rs("SITE")
                        Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
                        SELMA301RS.MoreResults
                        SELMA301RS.Requery
                
                    End If
                    If SELMA301RS.RowCount > 0 Then
                        If Trim(SELMA301RS("SITEDESC")) = "site not specified/recorded" Then
                            cmb_sites.Text = ""
                        Else
                            cmb_sites.Text = SELMA301RS("SITEDESC")
                        End If
                        Else
'                        venkat*********
'                            Dim viewsitedupqr As rdoQuery
'                            Dim viewsiteduprs As rdoResultset
'                            Dim viewstr As String
'                            viewstr = "select sitecd,sitedesc from dislib.sitedup where sitecd = ?"
'                            Set viewsitedupqr = rdb.CreateQuery(viewsitedupqr1, viewstr)
'                            viewsitedupqr(0) = seldrequisitn_rs("SITE")
'                            Set viewsiteduprs = viewsitedupqr.OpenResultset(1, 3)
'                            cmb_sites.Text = viewsiteduprs("SITEDESC")
                    End If
                    
                    cmb_laterality1.Text = seldrequisitn_rs("LATERALITS")
                    txt_metastasis.Text = seldrequisitn_rs("METASTASIS")
                    txt_outspecimen.Text = seldrequisitn_rs("OUTSSPECIM")
                    cmb_outblock.Text = seldrequisitn_rs("OUTBLOCK")
                    cmb_outslide.Text = seldrequisitn_rs("OUTSLIDES")
                    cmb_frozensec.Text = seldrequisitn_rs("FROZENSEC")
                    If IsNull(Trim(seldrequisitn_rs("FROZENDATE"))) Then
                    MSK_FROSECDT.Text = "__/__/____"
                    Else
                    MSK_FROSECDT.Text = seldrequisitn_rs("FROZENDATE")
                    End If
                    cmb_laterality = seldrequisitn_rs("LATERALITM")
                    TXT_OPPROCEDURE.Text = seldrequisitn_rs("OPPROCEDUR")
                    'CMB_PREINVEST.Text = Trim(SELRREQUISITN_RS2("PREVEXAM"))
                    cmb_refdoc.Text = Trim(seldrequisitn_rs("ReFDOC"))
                    cmb_reftodoc.Text = Trim(seldrequisitn_rs("REFTODOC"))
                    If seldrequisitn_rs.EOF Then
                        cmb_spfnac.Text = " " & Trim(seldrequisitn_rs("LABNO"))
                    Else
                        cmb_spfnac.Text = ""
                    End If
                    If seldrequisitn_rs.EOF = False Then
                        txt_nmeterial.Text = " " & Trim(seldrequisitn_rs("NMETERIALS"))
                    Else
                        txt_nmeterial.Text = ""
                    End If
                    If Len(Trim(seldrequisitn_rs("DURILLNESS"))) > 1 Then
                    
                    I = InStr(seldrequisitn_rs("DURILLNESS"), "/")
                    txt_year.Text = Mid(seldrequisitn_rs("DURILLNESS"), 1, I - 1)
                    CMB_DURILLMONTH.Text = Mid(seldrequisitn_rs("DURILLNESS"), I + 1, 2)
                    
                    End If
                    SPENABLE ("DIS")
                    lbl_spreqno1.Visible = False
                    lbl_sppathno.Visible = True
                    txt_sppathno.Visible = False
                    txt_spreqno.Visible = False
                    ' Previously Investigated requisition nos
'                    If SELRREQFLAG = False Then
'                         SELRREQ_QR(0) = txt_dept.Text
'                         SELRREQ_QR(1) = txt_modcode.Text
'                         SELRREQ_QR(2) = Msk_caseno.FormattedText
'                         Set SELRREQ_RS = SELRREQ_QR.OpenResultset(1, 3)
'                         SELRREQFLAG = True
'                    Else
'                         SELRREQ_QR(0) = txt_dept.Text
'                         SELRREQ_QR(1) = txt_modcode.Text
'                         SELRREQ_QR(2) = Msk_caseno.FormattedText
'                         SELRREQ_RS.MoreResults
'                         SELRREQ_RS.Requery
'
'                     End If
'                     If SELRREQ_RS.RowCount > 0 Then
'                         cmb_preinvest.clear
'                         While Not SELRREQ_RS.EOF
'                             cmb_preinvest.AddItem SELRREQ_RS("REQNNO")
'                             SELRREQ_RS.MoveNext
'                         Wend
'                         cmb_preinvest.Enabled = True
'                     End If
''commented by mono
'                    If seLdreqrsflag = False Then
'                        SELDREQQR(0) = Msk_caseno.FormattedText
'                        Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
'                        seLdreqrsflag = True
'                    Else
'                        SELDREQQR(0) = Msk_caseno.FormattedText
'                        SELDREQRS.MoreResults
'                        SELDREQRS.Requery
'                    End If
'                If SELDREQRS.RowCount > 0 Then
'                    cmb_preinvest.clear
'                    While Not SELDREQRS.EOF
'                        cmb_preinvest.AddItem SELDREQRS("PATHNO") & "$" & SELDREQRS("REQNO")
'                        SELDREQRS.MoveNext
'                    Wend
'                    cmb_preinvest.ListIndex = 0
'                End If
                    ' Previously Investigated requisition nos
                    
                    
                    
                End If
End Sub

Public Sub IHVIEW()
''                immunohistochemistry.Visible = True
''                lbl_pathno.Visible = True
''                txt_pathno.Visible = True
''                cmb_ihprocode.Text = SELRREQUISITN_RS2("SER_CODE")
'''                lbl_ihpathno.Visible = False
'''                txt_ihpathno.Visible = False
''
''                If SELMA101FLAG = False Then
''                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS2("SER_CODE"))
''                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
''                    SELMA101FLAG = True
''                Else
''                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS2("SER_CODE"))
''                    SELMA101_RS.MoreResults
''                    SELMA101_RS.Requery
''                End If
''                If SELMA101_RS.RowCount > 0 Then
''                    cmb_ihprocedure.Text = SELMA101_RS("SER_DET")
''                End If
''                If SELDREQUISITNFLAG = False Then
''                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
''                    Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
''                    SELDREQUISITNFLAG = True
''                Else
''                    seldrequisitn_qr(0) = req_chglen(Trim(cmb_reqno.Text))
''                    seldrequisitn_rs.MoreResults
''                    seldrequisitn_rs.Requery
''                End If
''                txt_pathno.Text = seldrequisitn_rs("pathno")
''                cmb_Ihsoutsideblock = seldrequisitn_rs("OUTBLOCK")
''                cmb_Ihsoutslide.Text = seldrequisitn_rs("OUTSLIDES")
''                If Len(Trim(seldrequisitn_rs("DURILLNESS"))) > 1 Then
''
''                    i = InStr(seldrequisitn_rs("DURILLNESS"), "/")
''                    txt_year.Text = Mid(seldrequisitn_rs("DURILLNESS"), 1, i - 1)
''                    CMB_DURILLMONTH.Text = Mid(seldrequisitn_rs("DURILLNESS"), i + 1, 2)
''                End If
''                'cmb_Ihspreinv.Text = Trim(SELRREQUISITN_RS2("PREVEXAM"))
''                cmb_refdoc.Text = Trim(seldrequisitn_rs("ReFDOC"))
''                cmb_reftodoc.Text = Trim(seldrequisitn_rs("REFTODOC"))
''                If seldrequisitn_rs.EOF = False Then
''                    txt_ihmeterial.Text = " " & seldrequisitn_rs("NMETERIALS").Value
''                Else
''                    txt_ihmeterial.Text = ""
''                End If
''                IHENABL ("DIS")
''                lbl_ihreqno.Visible = False
''                lbl_ihpathno.Visible = False
''                txt_ihreqno.Visible = False
''                txt_ihpathno.Visible = False
''                ' Previously Investigated requisition nos
''                    If SELRREQFLAG = False Then
''                         SELRREQ_QR(0) = txt_dept.Text
''                         SELRREQ_QR(1) = txt_modcode.Text
''                         SELRREQ_QR(2) = Msk_caseno.FormattedText
''                         Set SELRREQ_RS = SELRREQ_QR.OpenResultset(1, 3)
''                         SELRREQFLAG = True
''                    Else
''                         SELRREQ_QR(0) = txt_dept.Text
''                         SELRREQ_QR(1) = txt_modcode.Text
''                         SELRREQ_QR(2) = Msk_caseno.FormattedText
''                         SELRREQ_RS.MoreResults
''                         SELRREQ_RS.Requery
''
''                     End If
''                     If SELRREQ_RS.RowCount > 0 Then
''                         cmb_preinvest.clear
''                         While Not SELRREQ_RS.EOF
''                             cmb_Ihspreinv.AddItem SELRREQ_RS("REQNNO")
''                             SELRREQ_RS.MoveNext
''                         Wend
''                         cmb_Ihspreinv.Enabled = True
''
''                     End If
''                     If SELRREQUISITN_RS2.RowCount > 0 Then
''                                If SELRREQUISITN_RS2("RECORDNO") = 1 Then
''                                CMB_IHNIP.Text = "VPP"
''                                End If
''                                If SELRREQUISITN_RS2("RECORDNO") = 2 Then
''                                CMB_IHNIP.Text = "LR"
''                                End If
''                                If SELRREQUISITN_RS2("RECORDNO") = 3 Then
''                                CMB_IHNIP.Text = "BAR"
''                                End If
''                                If SELRREQUISITN_RS2("RECORDNO") = 4 Then
''                                CMB_IHNIP.Text = "DrOp"
''                                End If
''                            End If
''                    ' Previously Investigated requisition nos
''
''
''
''
End Sub
Public Sub NO_CHARGESP()

'*****
On Error GoTo MSG:
No_Chargeflag = False
 
If Left(Msk_caseno.FormattedText, 1) = "P" And Trim(txt_dept.Text) = "FZZ" Then
  'pre_sercodech(0) = Trim(cbo_proccode.Text)
  'Set selma101_RS = pre_sercodech.OpenResultset(1, 1)
      'If selma101_RS.RowCount = 1 And selma101_RS("yn_flag") = "Y" Then
          PRE_QUERY(0) = Msk_caseno.FormattedText
          'PRE_QUERY(1) = selma101_RS("FROM_DATE")
          'PRE_QUERY(2) = selma101_RS("TO_DATE")
          Set PREV_RS = PRE_QUERY.OpenResultset(1, 3)
            If PREV_RS.RowCount > 0 And Trim(txt_dept.Text) = "FZZ" Then
                   No_Chargeflag = True
                   NoMemo = True
              Else
                  If Trim(txt_catgory.Text) = "NC" Then
                    No_Chargeflag = True
                    NoMemo = True
                  Else
                        No_Chargeflag = False
                        NoMemo = False
'''                       MsgBox "Please Prepare the Memo !!", vbInformation, "Requisition"
                  End If
            End If
   
Else
    If SELMA101FLAG = False Then
        SELMA101_QR(0) = Trim(cmb_procedurecode.Text)
        Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
        SELMA101FLAG = True
     Else
        SELMA101_QR(0) = Trim(cmb_procedurecode.Text)
        SELMA101_RS.MoreResults
        SELMA101_RS.Requery
    End If
'  SELMA101_QR(0) = Trim(cbo_proccode.Text)
'  SELMA101_RS.MoreResults
'  SELMA101_RS.Requery

    If SELMA101_RS.RowCount > 0 Then
        If Trim(txt_catgory.Text) = "C" Then
            If SELMA101_RS("C_CH") = 0 Then
                No_Chargeflag = True
            End If
        ElseIf Trim(txt_catgory.Text) = "NC" Then ' Or Opt_S.Value = True Then
            If SELMA101_RS("NC_CH") = 0 Then
                No_Chargeflag = True
            End If
        ElseIf opt_staff.Value = True Then
                No_Chargeflag = True
        End If
    Else
        'MsgBox "There is NO  RECORD with this SERVICE CODE."   'COMMENTED BY SHWETA 18-AUG-09
    End If
End If

If No_Chargeflag = False Then
        NOCHARG_QR(0) = Msk_caseno.FormattedText
         If NOCHARG_FLG = False Then
           Set NOCHARG_RS = NOCHARG_QR.OpenResultset(2, 3)
           NOCHARG_FLG = True
         Else
           NOCHARG_RS.MoreResults
           NOCHARG_RS.Requery
         End If
          If NOCHARG_RS.RowCount > 0 Then
            'CHCKING FREE CODE FOR PREVENTIVE ONCOLOGY
             PRFREE_QR(0) = Trim(cmb_procedurecode.Text)
               If PRFREE_FLG = False Then
                  Set PRFREE_RS = PRFREE_QR.OpenResultset(2, 3)
                  PRFREE_FLG = True
               Else
                  PRFREE_RS.MoreResults
                  PRFREE_RS.Requery
               End If
               If PRFREE_RS.RowCount > 0 Then
                  No_Chargeflag = True
                  NoMemo = True
               End If
          Else
             No_Chargeflag = False
             NoMemo = False
            MsgBox "Please Prepare the Memo !!", vbInformation, "Requisition"
          End If
End If

    
Exit Sub
MSG:
MsgBox ERROR(Err)



'*****


On Error GoTo MSG:
''No_Chargeflag = False
''
''SELMA101_QR(0) = Trim(cbo_proccode.Text)
''SELMA101_RS.MoreResults
''SELMA101_RS.Requery
''If SELMA101_RS.RowCount > 0 Then
''    If Trim(txt_catgory.Text) = "C" Then
''        If SELMA101_RS("C_CH") = 0 Then
''            No_Chargeflag = True
''        End If
''    ElseIf Trim(txt_catgory.Text) = "NC" Then
''        If SELMA101_RS("NC_CH") = 0 Then
''            No_Chargeflag = True
''        End If
''    End If
''Else
''    MsgBox "There is NO  RECORD with this SERVICE CODE."
''End If
''Exit Sub
''MSG:
''MsgBox Error(Err)
End Sub
Public Sub NO_CHARGEFA1()
'*****
On Error GoTo MSG:
No_Chargeflag = False
 
''If Left(CASENUM, 1) = "P" And Trim(lbl_deptcode.Caption) = "IZZ" Then
''  'pre_sercodech(0) = Trim(cbo_proccode.Text)
''  'Set selma101_RS = pre_sercodech.OpenResultset(1, 1)
''      'If selma101_RS.RowCount = 1 And selma101_RS("yn_flag") = "Y" Then
''          PRE_QUERY(0) = CASENUM
''          'PRE_QUERY(1) = selma101_RS("FROM_DATE")
''          'PRE_QUERY(2) = selma101_RS("TO_DATE")
''          Set PREV_RS = PRE_QUERY.OpenResultset(1, 1)
''            If PREV_RS.RowCount > 0 And Trim(lbl_deptcode.Caption) = "IZZ" Then
''                   No_Chargeflag = True
''              Else
''                   No_Chargeflag = False
''            End If
''      'Else
''
''      '    No_Chargeflag = False
''
''      'End If
''Else
'  SELMA101_QR(0) = Trim(cbo_proccode.Text)
'  SELMA101_RS.MoreResults
'  SELMA101_RS.Requery
'
'    If SELMA101_RS.RowCount > 0 Then
'        If Trim(lbl_catg.Caption) = "C" Then
'            If SELMA101_RS("C_CH") = 0 Then
'                No_Chargeflag = True
'            End If
'        ElseIf Trim(lbl_catg.Caption) = "NC" Then ' Or Opt_S.Value = True Then
'            If SELMA101_RS("NC_CH") = 0 Then
'                No_Chargeflag = True
'            End If
'        ElseIf Opt_S.Value = True Then
'                No_Chargeflag = True
'        End If
'    Else
'        MsgBox "There is NO  RECORD with this SERVICE CODE."
'    End If
''End If
'Exit Sub
'MSG:
'MsgBox Error(Err)



'*****
On Error GoTo MSG:
No_Chargeflag = False
If SELMA101FLAG = False Then
    SELMA101_QR(0) = Trim(cmb_fnprocode.Text)
    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
    SELMA101FLAG = True
 Else
    SELMA101_QR(0) = Trim(cmb_fnprocode.Text)
    SELMA101_RS.MoreResults
    SELMA101_RS.Requery
End If

'SELMA101_QR(0) = Trim(cmb_procedurecode.Text)
'SELMA101_RS.MoreResults
'SELMA101_RS.Requery
If SELMA101_RS.RowCount > 0 Then
    If Trim(txt_catgory.Text) = "C" Then
        If SELMA101_RS("C_CH") = 0 Then
            No_Chargeflag = True
        End If
    ElseIf Trim(txt_catgory.Text) = "NC" Then
        If SELMA101_RS("NC_CH") = 0 Then
            No_Chargeflag = True
        End If
    End If
    If Trim(txt_catgory.Text) = "B" Then
        If SELMA101_RS("B_HOSP_CH") = 0 Then
            No_Chargeflag = True
        End If
    End If
    
    NOCHARG_QR(0) = Msk_caseno.FormattedText
    If NOCHARG_FLG = False Then
      Set NOCHARG_RS = NOCHARG_QR.OpenResultset(2, 3)
      NOCHARG_FLG = True
    Else
      NOCHARG_RS.MoreResults
      NOCHARG_RS.Requery
    End If
    If NOCHARG_RS.RowCount > 0 Then
      'CHCKIN FREE CODE
        PRFREE_QR(0) = Trim(cmb_fnprocode.Text)
          If PRFREE_FLG = False Then
             Set PRFREE_RS = PRFREE_QR.OpenResultset(2, 3)
             PRFREE_FLG = True
          Else
             PRFREE_RS.MoreResults
             PRFREE_RS.Requery
          End If
          If PRFREE_RS.RowCount > 0 Then
             No_Chargeflag = True
             NoMemo = True
          End If
    End If
    
    
Else
    'MsgBox "There is NO  RECORD with this SERVICE CODE."  'COMMENTED BY SHWETA 18-AUG-09
End If
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub
'Private Sub MEMO_CHECK()
'Dim i
'Dim x
'i = Data1.Recordset.RecordCount
'If i <> 0 Then
'    Data1.Recordset.MoveFirst
'    Do While Not Data1.Recordset.EOF
'        If SELDMULPROCEFLAG = False Then
'             SELDMULPROCEQR(0) = req_chglen(Data1.Recordset("REQNO"))
'             Set SELDMULPROCERS = SELDMULPROCEQR.OpenResultset(1, 3)
'             SELDMULPROCEFLAG = True
'        Else
'             SELDMULPROCEQR(0) = req_chglen(Data1.Recordset("REQNO"))
'             SELDMULPROCERS.MoreResults
'             SELDMULPROCERS.Requery
'        End If
'        If Not SELDMULPROCERS.EOF Then
'            Do While Not SELDMULPROCERS.EOF
'                x = x + 1
'                SELDMULPROCERS.MoveNext
'            Loop
'        End If
'        Data1.Recordset.MoveNext
'    Loop
'End If
'If Val(x) > 11 Then
'    MsgBox "ONE MEMO CAN'T CANTAIN MORE THAN 11 PROCEDURES !!", vbInformation, "DIS"
'    V_PROCE = True
'    Exit Sub
'End If
'End Sub
Private Sub SPQUERY()
If seldrequisitn_flag10 = False Then
    seldrequisitn_qr10(0) = UCase(CASENO)
    Set seldrequisitn_rs10 = seldrequisitn_qr10.OpenResultset(1, 3)
    seldrequisitn_FLAG = True
Else
    seldrequisitn_qr10(0) = UCase(CASENO)
    seldrequisitn_rs10.MoreResults
    seldrequisitn_rs10.Requery
End If
If seldrequisitn_rs10.RowCount > 0 Then
    'prpathno = ""
    cmb_preinvest.clear
    If txt_modcode.Text = "CG" Or txt_modcode.Text = "CN" Then
        cmb_cpreinv.clear
    End If
    If txt_modcode.Text = "FA" Then
        cmb_fnpreinvest.clear
    End If
    While Not seldrequisitn_rs10.EOF
        'CMB_PREINVEST.AddItem seldrequisitn_rs10("PATHNO") & " " & "$" & " " & req_chglen(seldrequisitn_rs10("REQNNO"))
        cmb_preinvest.AddItem Trim(seldrequisitn_rs10("REQNNO")) & " $ " & Trim(seldrequisitn_rs10("PREPORTDT")) & " $ " & Trim(seldrequisitn_rs10("PATHNO"))
        cmb_cpreinv.AddItem Trim(seldrequisitn_rs10("REQNNO")) & " $ " & Trim(seldrequisitn_rs10("PREPORTDT")) & " $ " & Trim(seldrequisitn_rs10("PATHNO"))
                                '" $ " & Trim(MODDESC) &
        If txt_modcode.Text = "CG" Or txt_modcode.Text = "CN" Then
            cmb_cpreinv.AddItem Trim(seldrequisitn_rs10("REQNNO")) & " $ " & Trim(seldrequisitn_rs10("PREPORTDT")) & " $ " & Trim(seldrequisitn_rs10("PATHNO"))
        End If
        If txt_modcode.Text = "FA" Then
            cmb_fnpreinvest.AddItem Trim(seldrequisitn_rs10("REQNNO")) & " $ " & Trim(seldrequisitn_rs10("PREPORTDT")) & " $ " & Trim(seldrequisitn_rs10("PATHNO"))
        End If
        seldrequisitn_rs10.MoveNext
        cmb_preinvest.Text = cmb_preinvest.List(0)
        If txt_modcode.Text = "CG" Or txt_modcode.Text = "CN" Then
            cmb_cpreinv.Text = cmb_cpreinv.List(0)
        End If
    Wend
End If
End Sub
Private Sub SPREPORT()
Dim str1 As String
Dim str2 As String
Dim STR3 As String
Dim I As Integer
On Error GoTo prtmsg:
DoEvents
Dim VAR_REQNNO As String
VAR_REQNNO = Trim(REQNO1)
If SELRREQNOFLAG = False Then
    SELRREQNOQR(0) = VAR_REQNNO
    Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
    SELRREQNOFLAG = True
Else
    SELRREQNOQR(0) = VAR_REQNNO
    SELRREQNORS.MoreResults
    SELRREQNORS.Requery
End If
'seldrequisitn_qr10.QueryTimeout = 150
If seldrequisitn_flag11 = False Then
    seldrequisitn_qr11(0) = Msk_caseno.FormattedText
    Set seldrequisitn_rs11 = seldrequisitn_qr11.OpenResultset(1, 3)
    seldrequisitn_flag11 = True
Else
    seldrequisitn_qr11(0) = Msk_caseno.FormattedText
    seldrequisitn_rs11.MoreResults
    seldrequisitn_rs11.Requery
End If
If seldrequisitn_rs11.RowCount > 0 Then
    prpathno = ""
    'CMB_PATHNO.clear
    While Not seldrequisitn_rs11.EOF
        prpathno = prpathno & " " & seldrequisitn_rs11("pathno")
        'CMB_PATHNO.AddItem seldrequisitn_rs10("PATHNO") & "$" & seldrequisitn_rs10("REQNO")
        seldrequisitn_rs11.MoveNext
    Wend
    'txt_prepathreqno.Text = prpathno
End If
If SELDADDENDUM_FLAG = False Then
    SELDADDENDUM_QR(0) = VAR_REQNNO
    Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
    SELDADDENDUM_FLAG = True
Else
    SELDADDENDUM_QR(0) = VAR_REQNNO
    SELDADDENDUM_RS.MoreResults
    SELDADDENDUM_RS.Requery
End If

For I = 0 To 25
    CrystalReport1.Formulas(I) = ""
Next I


str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
CrystalReport1.Formulas(0) = "REQNO = '" & Trim(REQNO1) & "'"
If opt_patient.Value = True Then
        If SELRREQNORS("COMMITFLAG") = "Y" Then
            If SELDADDENDUM_RS.RowCount > 0 Then
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZENADD.RPT"
            Else
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZEN.RPT"
            End If
        Else
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZEN.RPT"
        End If
        SELPATHDREQUISITN_QR(0) = VAR_REQNNO
        Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
        If selpathdrequisitn_rs.RowCount > 0 Then
             'CrystalReport1.Formulas(1) = "PREPATH = 'Prev. Path No." & Trim(prpathno) & "'"
             CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
        Else
        
             CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
        End If
        selta1012_qr(0) = Msk_caseno.FormattedText
        Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
        If selta1012_rs.RowCount > 0 Then
            CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
        Else
            CrystalReport1.Formulas(2) = "PROVDIA = '" & "" & "'"
        End If
        
        Dim z
        Dim Y
        Dim w
        Dim u
        
        SELMA301QR(0) = selpathdrequisitn_rs("SITE")
        Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
        If SELMA301RS.RowCount > 0 Then
            CrystalReport1.Formulas(3) = "SIT = '" & Trim(SELMA301RS("SITEDESC")) & "'"
        Else
            CrystalReport1.Formulas(3) = "SIT = '" & "" & "'"
        End If
        If selpathdrequisitn_rs.RowCount > 0 Then
            If Trim(selpathdrequisitn_rs("OPPROCEDUR")) <> "" Then
                z = InStr(1, Trim(selpathdrequisitn_rs("OPPROCEDUR")), "'")
                If z <> 0 Then
                    Y = Trim(selpathdrequisitn_rs("OPPROCEDUR"))
                    u = Len(Y)
                    X = Left(Y, z - 1)
                    w = Right(Y, u - (z))
                    Y = X & w
                    CrystalReport1.Formulas(4) = "OPPRO = '" & Y & "'"
                Else
                    CrystalReport1.Formulas(4) = "OPPRO = '" & Trim(selpathdrequisitn_rs("OPPROCEDUR")) & "'"
                End If
            Else
                CrystalReport1.Formulas(4) = "OPPRO = '" & Trim(selpathdrequisitn_rs("OPPROCEDUR")) & "'"
            End If
        Else
            CrystalReport1.Formulas(4) = "OPPRO = '" & "" & "'"
        End If
        str1 = ""
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
            str1 = "Entered By " & Trim(SELRREQNORS("enterby"))
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        SELDFREPORT_QR(0) = VAR_REQNNO
        Set SELDFREPORT_RS = SELDFREPORT_QR.OpenResultset(1, 3)
        If SELDFREPORT_RS.RowCount > 0 Then
            If Len(Trim(SELDFREPORT_RS("FSPREPORT"))) > 1 Then
                CrystalReport1.Formulas(6) = "FRSEC = '" & "Frozen Section" & "'"
            Else
                CrystalReport1.Formulas(6) = "FRSEC = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("IHREPORT"))) > 1 Then
                CrystalReport1.Formulas(7) = "IHREP = '" & "Immunohistochemistry" & "'"
            Else
                CrystalReport1.Formulas(7) = "IHREP = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("grossdesc"))) > 1 Then
                CrystalReport1.Formulas(10) = "GROSDESC = '" & "Grossdescription" & "'"
            Else
                CrystalReport1.Formulas(10) = "GROSDESC = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("spreport"))) > 1 Then
                CrystalReport1.Formulas(11) = "MICRODESC = '" & "Microscopic Description" & "'"
            Else
                CrystalReport1.Formulas(11) = "MICRODESC = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("Impression"))) > 2 Then
                CrystalReport1.Formulas(12) = "IMPN = '" & "Impression" & "'"
            Else
                CrystalReport1.Formulas(12) = "IMPN = '" & "" & "'"
            End If
        End If
        
        CrystalReport1.Formulas(8) = "USRID = '" & str & "'"
        
        'CrystalReport1.Formulas(10) = "GROSDESC = '" & "Gross description" & "'"
        'CrystalReport1.Formulas(11) = "MICRODESC = '" & "Microscopic Description" & "'"
        CrystalReport1.WindowState = crptMaximized
      '  CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
    Else
        SELEMPBASINFO_QR(0) = MaskEdBox1.Text
        SELEMPBASINFO_RS.MoreResults
        SELEMPBASINFO_RS.Requery
'        If SELEMPBASINFO_RS.RowCount > 0 Then
'            txt_sex.Text = SELEMPBASINFO_RS("GENDER")
'            txt_name.Text = Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME"))
'            txt_age = DateDiff("YYYY", SELEMPBASINFO_RS("DOB"), CStr(DATE11))
'        End If

        If SELRREQNORS("COMMITFLAG") = "Y" Then
            If SELDADDENDUM_RS.RowCount > 0 Then
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZENADDSTAFF.RPT"
            Else
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZENSTAFF.RPT"
            End If
        Else
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZENSTAFF.RPT"
        End If
        SELPATHDREQUISITN_QR(0) = VAR_REQNNO
        Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
        If selpathdrequisitn_rs.RowCount > 0 Then
             CrystalReport1.Formulas(1) = "PREPATH = '" & Trim(selpathdrequisitn_rs("prepathno")) & "'"
        Else
            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
        End If
        selta1012_qr(0) = MaskEdBox1.FormattedText
        Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
        If selta1012_rs.RowCount > 0 Then
            CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
        Else
            CrystalReport1.Formulas(2) = "PROVDIA = '" & "" & "'"
        End If
        
        
        SELMA301QR(0) = selpathdrequisitn_rs("SITE")
        Set SELMA301RS = SELMA301_QR.OpenResultset(1, 3)
        If SELMA301RS.RowCount > 0 Then
            CrystalReport1.Formulas(3) = "SIT = '" & Trim(SELMA301RS("SITEDESC")) & "'"
        Else
            CrystalReport1.Formulas(3) = "SIT = '" & "" & "'"
        End If
        If selpathdrequisitn_rs.RowCount > 0 Then
            CrystalReport1.Formulas(4) = "OPPRO = '" & Trim(selpathdrequisitn_rs("OPPROCEDUR")) & "'"
        Else
            CrystalReport1.Formulas(4) = "OPPRO = '" & "" & "'"
        End If
        str1 = ""
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
            str1 = "Entered By " & Trim(SELRREQNORS("enterby"))
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        SELDFREPORT_QR(0) = VAR_REQNNO
        Set SELDFREPORT_RS = SELDFREPORT_QR.OpenResultset(1, 3)
        If SELDFREPORT_RS.RowCount > 0 Then
            If Len(Trim(SELDFREPORT_RS("FSPREPORT"))) > 1 Then
                CrystalReport1.Formulas(6) = "FRSEC = '" & "Frozen Section" & "'"
            Else
                CrystalReport1.Formulas(6) = "FRSEC = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("IHREPORT"))) > 1 Then
                CrystalReport1.Formulas(7) = "IHREP = '" & "Immunohistochemistry" & "'"
            Else
                CrystalReport1.Formulas(7) = "IHREP = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("grossdesc"))) > 1 Then
                CrystalReport1.Formulas(10) = "GROSDESC = '" & "Grossdescription" & "'"
            Else
                CrystalReport1.Formulas(10) = "GROSDESC = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("spreport"))) > 1 Then
                CrystalReport1.Formulas(11) = "MICRODESC = '" & "Microscopic Description" & "'"
            Else
                CrystalReport1.Formulas(11) = "MICRODESC = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("Impression"))) > 2 Then
                CrystalReport1.Formulas(12) = "IMPN = '" & "Impression" & "'"
            Else
                CrystalReport1.Formulas(12) = "IMPN = '" & "" & "'"
            End If
         
        End If
        CrystalReport1.Formulas(8) = "USRID = '" & str & "'"
        CrystalReport1.Formulas(13) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
        CrystalReport1.Formulas(14) = "DOB = '" & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREQNORS("REQNDT")) & "'"
        CrystalReport1.Formulas(15) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
        CrystalReport1.WindowState = crptMaximized
        'CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
    End If
'CrystalReport1.WindowState = crptMaximized
'CrystalReport1.Action = 1
Exit Sub
prtmsg: MsgBox ERROR(Err)
'txt_msg.Text = ""
'btnenabl ("print")
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub
Private Sub FAREPORT()
Dim VAR_REQNNO As String
VAR_REQNNO = Trim(REQNO1)
If SELRREQNOFLAG = False Then
    SELRREQNOQR(0) = VAR_REQNNO
    Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
    SELRREQNOFLAG = True
Else
    SELRREQNOQR(0) = VAR_REQNNO
    SELRREQNORS.MoreResults
    SELRREQNORS.Requery
End If
'seldrequisitn_qr10.QueryTimeout = 300
'If seldrequisitn_flag10 = False Then
'    seldrequisitn_qr10(0) = MaskEdBox1.FormattedText
'    Set seldrequisitn_rs10 = seldrequisitn_qr10.OpenResultset(1, 3)
'    seldrequisitn_flag10 = True
'Else
'    seldrequisitn_qr10(0) = MaskEdBox1.FormattedText
'    seldrequisitn_rs10.MoreResults
'    seldrequisitn_rs10.Requery
'End If
'If seldrequisitn_rs10.RowCount > 0 Then
'    prpathno = ""
'    CMB_PATHNO.clear
'    While Not seldrequisitn_rs10.EOF
'        prpathno = prpathno & " " & seldrequisitn_rs10("pathno")
'        CMB_PATHNO.AddItem seldrequisitn_rs10("PATHNO") & "$" & seldrequisitn_rs10("REQNO")
'        seldrequisitn_rs10.MoveNext
'    Wend
'    txt_prepathreqno.Text = prpathno
'End If
'If SELDADDENDUM_FLAG = False Then
'    SELDADDENDUM_QR(0) = VAR_REQNNO
'    Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
'    SELDADDENDUM_FLAG = True
'Else
'    SELDADDENDUM_QR(0) = VAR_REQNNO
'    SELDADDENDUM_RS.MoreResults
'    SELDADDENDUM_RS.Requery
'End If

For I = 0 To 25
    CrystalReport1.Formulas(I) = ""
Next I

str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
'CRYSTALREPORT1.Formulas(0) = "REQNO = '" & Trim(REQNO1) & "'"

If opt_patient.Value = True Then
        If SELRREQNORS("REQNSTAT") <> "F" Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFA.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
                CrystalReport1.Formulas(1) = "PROFIN = '" & "FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
                str1 = ""
                str1 = "Entered By " & SELRREQNORS("ENTERBY")
                CrystalReport1.Formulas(2) = "ENT = '" & str1 & "'"
             Else
    
                CrystalReport1.Formulas(1) = "PROFIN = '" & "FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
                CrystalReport1.Formulas(2) = "ENT = '" & str1 & "'"
            End If
            CrystalReport1.Formulas(3) = "USRID = '" & str & "'"
        Else
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFAADD.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            CrystalReport1.Formulas(3) = "USRID = '" & str & "'"
        End If
        'CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
    Else
        SELEMPBASINFO_QR(0) = MaskEdBox1.Text
        SELEMPBASINFO_RS.MoreResults
        SELEMPBASINFO_RS.Requery
        If SELRREQNORS("REQNSTAT") <> "F" Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFASTAFF.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
                CrystalReport1.Formulas(5) = "PROFIN = '" & "FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
                str1 = ""
                str1 = "Entered By " & SELRREQNORS("ENTERBY")
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
             Else
                CrystalReport1.Formulas(5) = "PROFIN = '" & "FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            End If
            CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
        Else
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFAADDSTAFF.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
        End If
        CrystalReport1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
        CrystalReport1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
        CrystalReport1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREQNORS("REQNDT"))
        CrystalReport1.WindowState = crptMaximized
        'CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
    End If
End Sub
Private Sub CGCNREPORT()
Dim VAR_REQNNO As String
VAR_REQNNO = Trim(REQNO1)
If SELRREQNOFLAG = False Then
    SELRREQNOQR(0) = VAR_REQNNO
    Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
    SELRREQNOFLAG = True
Else
    SELRREQNOQR(0) = VAR_REQNNO
    SELRREQNORS.MoreResults
    SELRREQNORS.Requery
End If
'seldrequisitn_qr10.QueryTimeout = 300
'If seldrequisitn_flag10 = False Then
'    seldrequisitn_qr10(0) = MaskEdBox1.FormattedText
'    Set seldrequisitn_rs10 = seldrequisitn_qr10.OpenResultset(1, 3)
'    seldrequisitn_flag10 = True
'Else
'    seldrequisitn_qr10(0) = MaskEdBox1.FormattedText
'    seldrequisitn_rs10.MoreResults
'    seldrequisitn_rs10.Requery
'End If
'SELPATHDREQUISITN_QR(0) = VAR_REQNNO
'Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
'
'If seldrequisitn_rs10.RowCount > 0 Then
'    prpathno = ""
'    CMB_PATHNO.clear
'    While Not seldrequisitn_rs10.EOF
'        prpathno = prpathno & " " & seldrequisitn_rs10("pathno")
'        CMB_PATHNO.AddItem seldrequisitn_rs10("PATHNO") & "$" & seldrequisitn_rs10("REQNO")
'        seldrequisitn_rs10.MoveNext
'    Wend
'    txt_prepathreqno.Text = prpathno
'End If
'If SELDADDENDUM_FLAG = False Then
'    SELDADDENDUM_QR(0) = VAR_REQNNO
'    Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
'    SELDADDENDUM_FLAG = True
'Else
'    SELDADDENDUM_QR(0) = VAR_REQNNO
'    SELDADDENDUM_RS.MoreResults
'    SELDADDENDUM_RS.Requery
'End If

'SELDADDENDUM_QR(0) = VAR_REQNNO
'SELDADDENDUM_RS.MoreResults
'SELDADDENDUM_RS.Requery

For I = 0 To 25
    CrystalReport1.Formulas(I) = ""
Next I

str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"

If opt_patient.Value = True Then
    If SELRREQNORS("REQNSTAT") <> "F" Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCG.RPT"
        'CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "CYTOLOGY REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(SELRREQNORS("ENTERBY"))
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "CYTOLOGY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        SELPATHDREQUISITN_QR(0) = VAR_REQNNO
        Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
        CrystalReport1.Formulas(12) = "NOSLIDES=" & "" & selpathdrequisitn_rs("LABNO")
        If Trim(SELRREQNORS("PREPORTBY")) <> "" And Trim(SELRREQNORS("FREPORTBY")) <> "" And UCase(Trim(SELRREQNORS("PREPORTBY"))) <> UCase(Trim(SELRREQNORS("FREPORTBY"))) Then
            CrystalReport1.Formulas(11) = "P= '" & Trim(SELRREQNORS("PREPORTBY")) & "'"
            CrystalReport1.Formulas(12) = "F= '" & Trim(SELRREQNORS("FREPORTBY")) & "'"
            CrystalReport1.Formulas(13) = "PCAP= '" & "Cytotechnologist" & " '"
            CrystalReport1.Formulas(14) = "FCAP= '" & "Pathologist" & " '"
        End If
        If Trim(SELRREQNORS("PREPORTBY")) <> "" And Trim(SELRREQNORS("FREPORTBY")) = "" Then
            CrystalReport1.Formulas(11) = "PONLY= '" & Trim(SELRREQNORS("PREPORTBY")) & "'"
            'CrystalReport1.Formulas(12) = "F= '" & "" & "'"
            CrystalReport1.Formulas(13) = "PFCAP= '" & "Cytotechnologist / Pathologist" & " '"
        End If
        If Trim(SELRREQNORS("PREPORTBY")) <> "" And Trim(SELRREQNORS("FREPORTBY")) <> "" And UCase(Trim(SELRREQNORS("PREPORTBY"))) = UCase(Trim(SELRREQNORS("FREPORTBY"))) Then
            CrystalReport1.Formulas(11) = "PONLY= '" & Trim(SELRREQNORS("PREPORTBY")) & "'"
            CrystalReport1.Formulas(13) = "PFCAP= '" & "Cytotechnologist / Pathologist" & " '"
        End If
    Else
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCGADD.RPT"
        'CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
        CrystalReport1.Formulas(12) = "NOSLIDES = " & "" & selpathdrequisitn_rs("LABNO")
    End If
Else
    If SELRREQNORS("REQNSTAT") <> "F" Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCGSTAFF.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "CYTOLOGY REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(SELRREQNORS("ENTERBY"))
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "CYTOLOGY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
    Else
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCGADDSTAFF.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
    End If
    CrystalReport1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
    CrystalReport1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
    CrystalReport1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREQNORS("REQNDT"))
End If
CrystalReport1.WindowState = crptMaximized
'CRYSTALREPORT1.WindowState = crptMaximized
CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

CrystalReport1.Action = 1
End Sub
Private Sub RISREPORT()
Dim VAR_REQNNO As String
VAR_REQNNO = Trim(REQNO1)
If SELRREQNOFLAG = False Then
    SELRREQNOQR(0) = VAR_REQNNO
    Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
    SELRREQNOFLAG = True
Else
    SELRREQNOQR(0) = VAR_REQNNO
    SELRREQNORS.MoreResults
    SELRREQNORS.Requery
End If
'rq_grid_det(0) = Trim(VAR_REQNNO)
'rs_grid_det.MoreResults
'rs_grid_det.Requery

For I = 0 To 25
    CrystalReport1.Formulas(I) = ""
Next I

If SELRREQNORS.RowCount > 0 And (Trim(SELRREQNORS("REQNSTAT")) = "T" Or Trim(SELRREQNORS("REQNSTAT")) = "F") Then
    If opt_patient.Value = True Then
        CrystalReport1.Formulas(0) = "provtyp= '" & Trim(SELRREQNORS("ENTERBY")) & "'"
        'cr_rep.Formulas(1) = "reporttext= '" & Trim(txt_report.Text) & "'"
        'cr_rep.Formulas(0) = "impresstext= '" & Trim(txt_impression.Text) & "'"
        'cr_rep.Formulas(1) = "finalrep= '" & Trim(txt_finalreport.Text) & "'"
    
        SELRMODALITY_QR(0) = Mid(VAR_REQNNO, 4, 2)
        SELRMODALITY_RS.MoreResults
        SELRMODALITY_RS.Requery
        If SELRMODALITY_RS.RowCount > 0 Then
            CrystalReport1.Formulas(1) = "MODDES='" & SELRMODALITY_RS("MODLITYDES") & "'"
        Else
            CrystalReport1.Formulas(1) = "MODDES='" & "" & "'"
        End If
        CrystalReport1.Formulas(2) = "deptname= '" & "Department of Radiodiagnosis" & "'"
        CrystalReport1.SelectionFormula = "{RISREPORTS.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
        'cr_rep.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
        ''cr_rep.Formulas(0) = "provtypist = '" & txt_type1.Text & "'"
        str1 = ip
        str2 = DATE11
        STR3 = time1
        str1 = str1 & Space(5) & str2 & Space(5) & STR3
        str1 = Space(3) & str1
        CrystalReport1.Formulas(3) = "us_id = '" & str1 & "'"
        CrystalReport1.Formulas(4) = "INOUT = '" & Trim(SELRREQNORS("PATNSTAT")) & "'"
        CrystalReport1.Formulas(5) = "reqno = '" & req_chglen(VAR_REQNNO) & "'"
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Or Trim(SELRREQNORS("REQNSTAT")) = "F" Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\crpt_rep1.rpt"
        Else
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\crpt_rep2.rpt"
        End If
    Else
        SELEMPBASINFO_QR(0) = MaskEdBox1.Text
        SELEMPBASINFO_RS.MoreResults
        SELEMPBASINFO_RS.Requery
        
        CrystalReport1.Formulas(0) = "provtyp= '" & Trim(SELRREQNORS("ENTERBY")) & "'"
        SELRMODALITY_QR(0) = Mid(VAR_REQNNO, 4, 2)
        SELRMODALITY_RS.MoreResults
        SELRMODALITY_RS.Requery
        If SELRMODALITY_RS.RowCount > 0 Then
            CrystalReport1.Formulas(1) = "MODDES='" & SELRMODALITY_RS("MODLITYDES") & "'"
        Else
            CrystalReport1.Formulas(1) = "MODDES='" & "" & "'"
        End If
        CrystalReport1.Formulas(2) = "deptname= '" & "Department of Radiodiagnosis" & "'"
        CrystalReport1.SelectionFormula = "{RISREPORTS.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
        str1 = ip
        str2 = DATE11
        STR3 = time1
        str1 = str1 & Space(5) & str2 & Space(5) & STR3
        str1 = Space(3) & str1
        CrystalReport1.Formulas(3) = "us_id = '" & str1 & "'"
        CrystalReport1.Formulas(4) = "INOUT = '" & Trim(SELRREQNORS("PATNSTAT")) & "'"
        CrystalReport1.Formulas(5) = "reqno = '" & req_chglen(VAR_REQNNO) & "'"
        CrystalReport1.Formulas(6) = "PATNAME = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
        CrystalReport1.Formulas(7) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
        CrystalReport1.Formulas(5) = "PATAGE = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREQNORS("REQNDT"))
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Or Trim(SELRREQNORS("REQNSTAT")) = "F" Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\crpt_rep1STAFF.rpt"
        Else
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\crpt_rep2STAFF.rpt"
        End If
    End If
    'CRYSTALREPORT1.WindowState = crptMaximized
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
    CrystalReport1.Action = 1
    Exit Sub
Else
    Exit Sub
End If
End Sub

Public Sub MP()
 ' PROCEDURE CODE FILLING IN THE COMBOBOX
 
                If seLdreqrsmpflag = False Then
                    SELDREQMPQR(0) = Msk_caseno.FormattedText
                    Set SELDREQMPRS = SELDREQMPQR.OpenResultset(1, 3)
                    seLdreqrsmpflag = True
                Else
                    SELDREQMPQR(0) = Msk_caseno.FormattedText
                    SELDREQMPRS.MoreResults
                    SELDREQMPRS.Requery
                End If
                If SELDREQMPRS.RowCount > 0 Then
                    molecularpathology.cmb_prev.clear
                    While Not SELDREQMPRS.EOF
                        molecularpathology.cmb_prev.AddItem SELDREQMPRS("REQNO") & "$" & SELDREQMPRS("PATHNO")
                        SELDREQMPRS.MoveNext
                    Wend
                    molecularpathology.cmb_prev.ListIndex = 0
               End If
'               molecularpathology.list_proc.clear
 molecularpathology.show
' Load molecularpathology
 
End Sub

Public Sub MPASSIGN()
'*** INSERTING INT0 RREQUISITION
INSRREQUISITN_QR(0) = req_chglen(Trim(txt_bcreqno.Text))
INSRREQUISITN_QR(1) = GET_REVDATE(CStr(DATE11))
INSRREQUISITN_QR(2) = Trim(txt_time.Text)
INSRREQUISITN_QR(3) = Trim(Msk_caseno.FormattedText)
INSRREQUISITN_QR(4) = Trim(txt_dept.Text)
INSRREQUISITN_QR(5) = Trim(txt_modcode.Text)
INSRREQUISITN_QR(6) = Trim(txt_status.Text)
INSRREQUISITN_QR(7) = Trim(txt_catgory.Text)
INSRREQUISITN_QR(8) = "8888"
INSRREQUISITN_QR(9) = UCase(Trim(cmb_refdoc.Text))

'Added by Milind on 27-03-2006 for inserting date for the samples collected at outside
    If Check1.Value = 1 Then
        Dim Dt As String
        Dt = "18/02/1980"
        INSRREQUISITN_QR(15) = Format(CDate(Dt), "dd/mm/yyyy")
    Else
        INSRREQUISITN_QR(15) = Null
    End If
'upto here


    If SELRREQUISITNFLAG = False Then
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
        SELRREQUISITNFLAG = True
    Else
        SELRREQUISITN_QR(0) = Trim(Msk_caseno.FormattedText)
        SELRREQUISITN_RS.MoreResults
        SELRREQUISITN_RS.Requery
    End If
    If SELRREQUISITN_RS.RowCount > 0 Then
        INSRREQUISITN_QR(10) = "YES"
    Else
        INSRREQUISITN_QR(10) = "NO"
    End If
    
    If pathology.opt_patient.Value = True Then
       ''' Added by Sandhya dated 26.6.2009
    ''' for NC category & charged patient
       NO_CHARGE
       If txt_catgory.Text = "NC" And No_Chargeflag = True Then
            INSRREQUISITN_QR(11) = "NOCHARGE"
            INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
            INSRREQUISITN_QR(13) = "P"
            No_Chargeflag = False
       Else
            INSRREQUISITN_QR(11) = Null
            INSRREQUISITN_QR(12) = Null
            INSRREQUISITN_QR(13) = Null
        
       End If
    End If
    If pathology.opt_staff.Value = True Then
            INSRREQUISITN_QR(11) = "NOCHARGE"
            INSRREQUISITN_QR(12) = GET_REVDATE(CStr(DATE11))
            INSRREQUISITN_QR(13) = "P"
    End If
    INSRREQUISITN_QR(14) = 0
    INSRREQUISITN_QR.Execute
  
  'inserting into drequisitn
    
        INSDREQUISITN_QR(0) = req_chglen(Trim(txt_bcreqno.Text))
        INSDREQUISITN_QR(1) = " "
        'INSDREQUISITN_QR(1) = Trim(txt_sppathno.Text)
        INSDREQUISITN_QR(2) = Trim(Msk_caseno.FormattedText)
        INSDREQUISITN_QR(3) = Trim(StrSite)
        INSDREQUISITN_QR(4) = "" 'Trim(StrLaterality)
        INSDREQUISITN_QR(5) = Null
        INSDREQUISITN_QR(6) = Null
        INSDREQUISITN_QR(7) = Null
        INSDREQUISITN_QR(8) = Null
        INSDREQUISITN_QR(9) = Null
        INSDREQUISITN_QR(10) = Null
        INSDREQUISITN_QR(11) = Null
        INSDREQUISITN_QR(12) = Null
        INSDREQUISITN_QR(13) = Null
        INSDREQUISITN_QR(14) = Trim(cmb_refdoc.Text)
        INSDREQUISITN_QR(15) = Trim(cmb_reftodoc.Text)
        INSDREQUISITN_QR(16) = Trim(cmb_cyprocedure.Text)
        
        If Trim(pathology.txt_status.Text) = "IP" Then
            INSDREQUISITN_QR(17) = Trim(txt_ward.Text)
            INSDREQUISITN_QR(18) = Trim(txt_bedno.Text)
        Else
            INSDREQUISITN_QR(17) = Null
            INSDREQUISITN_QR(18) = Null
        End If
        INSDREQUISITN_QR(19) = " "
        INSDREQUISITN_QR(20) = Null
        INSDREQUISITN_QR(21) = Trim(ICU)
        INSDREQUISITN_QR.Execute
        
  '**** INSERTING INTO DREQUISITION
  'COMMENTED BY SHWETA 19-AUG
  '*** INSERTING INTO DMOLEPROC TABLE
      If CMB_TEXT.ListCount <> 0 Then
        For I = 0 To CMB_TEXT.ListCount - 1
            INSDMOLPROC_QR(0) = req_chglen(txt_bcreqno.Text)
            INSDMOLPROC_QR(1) = CMB_TEXT.List(I)
            INSDMOLPROC_QR(2) = GET_REVDATE(DATE11)
            INSDMOLPROC_QR(3) = Msk_caseno.FormattedText
            INSDMOLPROC_QR.Execute
        Next I
      End If

      If CMB_TEXT.ListCount <> 0 Then
       For I = 0 To CMB_TEXT.ListCount - 1
           INSDMULPROCE_QR(0) = req_chglen(txt_bcreqno.Text)
            INSDMULPROCE_QR(1) = CMB_TEXT.List(I)
            INSDMULPROCE_QR(2) = GET_REVDATE(DATE11)
            INSDMULPROCE_QR(3) = Msk_caseno.FormattedText
            INSDMULPROCE_QR.Execute
        Next I
     End If
 'COMMENTED BY SHWETA 19-AUG

' INSERTING INTO DMOLEREQ TABLE
  
  INSDMOLREQ_QR(0) = req_chglen(txt_bcreqno.Text)
  INSDMOLREQ_QR(1) = Msk_caseno.FormattedText
  INSDMOLREQ_QR(2) = Trim(StrSite)
  INSDMOLREQ_QR(3) = Trim(StrLaterality)
  INSDMOLREQ_QR(4) = Trim(StrFresh)
  INSDMOLREQ_QR(5) = Trim(StrArchive)
  INSDMOLREQ_QR.Execute

'  INSDMULPROCE_QR(0) = req_chglen(txt_bcreqno.Text)
'  INSDMULPROCE_QR(1) = msk_caseno.FormattedText
'  INSDMULPROCE_QR(2) = Trim(StrSite)
'  INSDMULPROCE_QR(3) = Trim(StrLaterality)
'  INSDMULPROCE_QR(4) = Trim(StrFresh)
'  INSDMULPROCE_QR(5) = Trim(StrArchive)
'  INSDMULPROCE_QR.Execute

  
  Data1.UpdateRecord
  Data1.Recordset.Bookmark = Data1.Recordset.LastModified
  
  Unload molecularpathology
End Sub

Public Sub MPPRINT()
                 CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DREQREPMP.RPT"
                 CrystalReport1.Formulas(0) = "REQNO = '" & Data1.Recordset.Fields("REQNO") & "'"
                    If txt_status = "OP" Then
                        CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & " Out Patient" & "'"
                    Else
                        CrystalReport1.Formulas(2) = "ST = '" & txt_catgory & "/" & "In Patient" & "'"
                    End If

                    If txt_status = "ST" Then
                        CrystalReport1.Formulas(2) = "ST = 'Out Patient'"
                    End If

                    CrystalReport1.Formulas(3) = "SEXAGE = '" & txt_sex & "/" & txt_Age & "'"
                    If Trim(txt_status.Text) = "IP" Then
                        CrystalReport1.Formulas(4) = "WRBED = '" & txt_ward & "/" & txt_bedno & "'"
                    Else
                        CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
                    End If
                If Mid(Trim(Data1.Recordset.Fields("REQNO")), 5, 2) = "BC" Then
                   CrystalReport1.Formulas(7) = "CAP = '" & "REQUISITION FOR BIOCHEMSITRY" & " '"
                End If
                CrystalReport1.Formulas(8) = "ENTERBY = '" & Trim(enter_by1) & "'"
                CrystalReport1.Formulas(9) = "mod = '" & Trim(txt_modcode) & "'"
                
End Sub

Public Sub NOCHARGEEL()
On Error GoTo MSG:
No_Chargeflag = False
 
If Left(Msk_caseno.FormattedText, 1) = "P" And Trim(txt_dept.Text) = "FEC" And Trim(txt_modcode.Text) = "EL" Then
  'pre_sercodech(0) = Trim(cbo_proccode.Text)
  'Set selma101_RS = pre_sercodech.OpenResultset(1, 1)
      'If selma101_RS.RowCount = 1 And selma101_RS("yn_flag") = "Y" Then
          PRE_QUERY(0) = Msk_caseno.FormattedText
          'PRE_QUERY(1) = selma101_RS("FROM_DATE")
          'PRE_QUERY(2) = selma101_RS("TO_DATE")
          Set PREV_RS = PRE_QUERY.OpenResultset(1, 3)
            If PREV_RS.RowCount > 0 And Trim(txt_dept.Text) = "FEC" Then
                   No_Chargeflag = True
                   NoMemo = True
              Else
                  If Trim(txt_catgory.Text) = "NC" Then
                    No_Chargeflag = True
                    NoMemo = True
                  Else
                        No_Chargeflag = False
                        NoMemo = False
'''                       MsgBox "Please Prepare the Memo !!", vbInformation, "Requisition"
                  End If
            End If
   
Else
    If SELMA101FLAG = False Then
        If CMB_TEXT.ListCount <> 0 Then
            For I = 0 To CMB_TEXT.ListCount - 1
                SELMA101_QR(0) = CMB_TEXT.List(I)
                Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                SELMA101FLAG = True
            Next I
        End If
     Else
        If CMB_TEXT.ListCount <> 0 Then
            For I = 0 To CMB_TEXT.ListCount - 1
                SELMA101_QR(0) = CMB_TEXT.List(I)  'Trim(cmb_procedurecode.Text)
                SELMA101_RS.MoreResults
                SELMA101_RS.Requery
            Next I
        End If
    End If
'  SELMA101_QR(0) = Trim(cbo_proccode.Text)
'  SELMA101_RS.MoreResults
'  SELMA101_RS.Requery

    If SELMA101_RS.RowCount > 0 Then
        If Trim(txt_catgory.Text) = "C" Then
            If SELMA101_RS("C_CH") = 0 Then
                No_Chargeflag = True
            End If
        ElseIf Trim(txt_catgory.Text) = "NC" Then ' Or Opt_S.Value = True Then
            If SELMA101_RS("NC_CH") = 0 Then
                No_Chargeflag = True
            End If
        ElseIf opt_staff.Value = True Then
                No_Chargeflag = True
        End If
    Else
        'MsgBox "There is NO  RECORD with this SERVICE CODE." 'COMMENTED BY SHWETA 18-AUG-09
    End If
End If

If No_Chargeflag = False Then
        NOCHARG_QR(0) = Msk_caseno.FormattedText
         If NOCHARG_FLG = False Then
           Set NOCHARG_RS = NOCHARG_QR.OpenResultset(2, 3)
           NOCHARG_FLG = True
         Else
           NOCHARG_RS.MoreResults
           NOCHARG_RS.Requery
         End If
          If NOCHARG_RS.RowCount > 0 Then
            'CHECKING FREE CODE FOR PREVENTIVE ONCOLOGY
            
             PRFREE_QR(0) = CMB_TEXT.List(I) 'Trim(cmb_procedurecode.Text)
               If PRFREE_FLG = False Then
                  Set PRFREE_RS = PRFREE_QR.OpenResultset(2, 3)
                  PRFREE_FLG = True
               Else
                  PRFREE_RS.MoreResults
                  PRFREE_RS.Requery
               End If
               If PRFREE_RS.RowCount > 0 Then
                  No_Chargeflag = True
                  NoMemo = True
               End If
          Else
             No_Chargeflag = False
             NoMemo = False
            MsgBox "Please Prepare the Memo !!", vbInformation, "Requisition"
          End If
End If

    
Exit Sub
MSG:
MsgBox ERROR(Err)



'*****


On Error GoTo MSG:
''No_Chargeflag = False
''
''SELMA101_QR(0) = Trim(cbo_proccode.Text)
''SELMA101_RS.MoreResults
''SELMA101_RS.Requery
''If SELMA101_RS.RowCount > 0 Then
''    If Trim(txt_catgory.Text) = "C" Then
''        If SELMA101_RS("C_CH") = 0 Then
''            No_Chargeflag = True
''        End If
''    ElseIf Trim(txt_catgory.Text) = "NC" Then
''        If SELMA101_RS("NC_CH") = 0 Then
''            No_Chargeflag = True
''        End If
''    End If
''Else
''    MsgBox "There is NO  RECORD with this SERVICE CODE."
''End If
''Exit Sub
''MSG:
''MsgBox Error(Err)

End Sub

Public Sub NOCHARGETM()
On Error GoTo MSG:
No_Chargeflag = False
 
If Left(Msk_caseno.FormattedText, 1) = "P" And Trim(txt_dept.Text) = "FEC" And Trim(txt_modcode.Text) = "TM" Then
  'pre_sercodech(0) = Trim(cbo_proccode.Text)
  'Set selma101_RS = pre_sercodech.OpenResultset(1, 1)
      'If selma101_RS.RowCount = 1 And selma101_RS("yn_flag") = "Y" Then
          PRE_QUERY(0) = Msk_caseno.FormattedText
          'PRE_QUERY(1) = selma101_RS("FROM_DATE")
          'PRE_QUERY(2) = selma101_RS("TO_DATE")
          Set PREV_RS = PRE_QUERY.OpenResultset(1, 3)
            If PREV_RS.RowCount > 0 And Trim(txt_dept.Text) = "FEC" Then
                   No_Chargeflag = True
                   NoMemo = True
              Else
                  If Trim(txt_catgory.Text) = "NC" Then
                    No_Chargeflag = True
                    NoMemo = True
                  Else
                        No_Chargeflag = False
                        NoMemo = False
'''                       MsgBox "Please Prepare the Memo !!", vbInformation, "Requisition"
                  End If
            End If
   
Else
    If SELMA101FLAG = False Then
        If CMB_TEXT.ListCount <> 0 Then
            For I = 0 To CMB_TEXT.ListCount - 1
                SELMA101_QR(0) = CMB_TEXT.List(I)
                Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                SELMA101FLAG = True
            Next I
        End If
     Else
        If CMB_TEXT.ListCount <> 0 Then
            For I = 0 To CMB_TEXT.ListCount - 1
                SELMA101_QR(0) = CMB_TEXT.List(I)  'Trim(cmb_procedurecode.Text)
                SELMA101_RS.MoreResults
                SELMA101_RS.Requery
            Next I
        End If
    End If
'  SELMA101_QR(0) = Trim(cbo_proccode.Text)
'  SELMA101_RS.MoreResults
'  SELMA101_RS.Requery

    If SELMA101_RS.RowCount > 0 Then
        If Trim(txt_catgory.Text) = "C" Then
            If SELMA101_RS("C_CH") = 0 Then
                No_Chargeflag = True
            End If
        ElseIf Trim(txt_catgory.Text) = "NC" Then ' Or Opt_S.Value = True Then
            If SELMA101_RS("NC_CH") = 0 Then
                No_Chargeflag = True
            End If
        ElseIf opt_staff.Value = True Then
                No_Chargeflag = True
        End If
    Else
        'MsgBox "There is NO  RECORD with this SERVICE CODE."  'COMMENTED BY SHWETA 18-AUG-09
    End If
End If

If No_Chargeflag = False Then
        NOCHARG_QR(0) = Msk_caseno.FormattedText
         If NOCHARG_FLG = False Then
           Set NOCHARG_RS = NOCHARG_QR.OpenResultset(2, 3)
           NOCHARG_FLG = True
         Else
           NOCHARG_RS.MoreResults
           NOCHARG_RS.Requery
         End If
          If NOCHARG_RS.RowCount > 0 Then
            'CHECKING FREE CODE FOR PREVENTIVE ONCOLOGY
            
             PRFREE_QR(0) = CMB_TEXT.List(I) 'Trim(cmb_procedurecode.Text)
               If PRFREE_FLG = False Then
                  Set PRFREE_RS = PRFREE_QR.OpenResultset(2, 3)
                  PRFREE_FLG = True
               Else
                  PRFREE_RS.MoreResults
                  PRFREE_RS.Requery
               End If
               If PRFREE_RS.RowCount > 0 Then
                  No_Chargeflag = True
                  NoMemo = True
               End If
          Else
             No_Chargeflag = False
             NoMemo = False
            MsgBox "Please Prepare the Memo !!", vbInformation, "Requisition"
          End If
End If

    
Exit Sub
MSG:
MsgBox ERROR(Err)



'*****


On Error GoTo MSG:
''No_Chargeflag = False
''
''SELMA101_QR(0) = Trim(cbo_proccode.Text)
''SELMA101_RS.MoreResults
''SELMA101_RS.Requery
''If SELMA101_RS.RowCount > 0 Then
''    If Trim(txt_catgory.Text) = "C" Then
''        If SELMA101_RS("C_CH") = 0 Then
''            No_Chargeflag = True
''        End If
''    ElseIf Trim(txt_catgory.Text) = "NC" Then
''        If SELMA101_RS("NC_CH") = 0 Then
''            No_Chargeflag = True
''        End If
''    End If
''Else
''    MsgBox "There is NO  RECORD with this SERVICE CODE."
''End If
''Exit Sub
''MSG:
''MsgBox Error(Err)

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





