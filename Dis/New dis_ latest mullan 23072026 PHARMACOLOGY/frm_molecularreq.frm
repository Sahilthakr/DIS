VERSION 5.00
Begin VB.Form molecularpathology 
   Caption         =   "Molecular Pathology"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11565
   LinkTopic       =   "Form2"
   ScaleHeight     =   7170
   ScaleWidth      =   11565
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ListBox list_code 
      Height          =   735
      ItemData        =   "frm_molecularreq.frx":0000
      Left            =   9060
      List            =   "frm_molecularreq.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   21
      Top             =   1950
      Width           =   1140
   End
   Begin VB.ListBox list_proc 
      Height          =   735
      Left            =   2550
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   1965
      Width           =   4785
   End
   Begin VB.CommandButton cmd_back 
      Caption         =   "&Back"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9105
      TabIndex        =   10
      Top             =   6285
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      Caption         =   "Archive"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1155
      Left            =   6585
      TabIndex        =   19
      Top             =   3705
      Width           =   3615
      Begin VB.CheckBox chk_paraffin 
         Caption         =   "Paraffin Block"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   345
         TabIndex        =   8
         Top             =   735
         Width           =   1485
      End
      Begin VB.CheckBox chk_atissue 
         Caption         =   "Tissue"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2220
         TabIndex        =   9
         Top             =   315
         Width           =   945
      End
      Begin VB.CheckBox chk_ablood 
         Caption         =   "Blood"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   345
         TabIndex        =   7
         Top             =   330
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fresh"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1155
      Left            =   2505
      TabIndex        =   18
      Top             =   3705
      Width           =   3690
      Begin VB.CheckBox chk_ftissue 
         Caption         =   "Tissue"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1995
         TabIndex        =   6
         Top             =   780
         Width           =   855
      End
      Begin VB.CheckBox chk_fblood 
         Caption         =   "Blood"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1980
         TabIndex        =   4
         Top             =   315
         Width           =   870
      End
      Begin VB.CheckBox chk_bonemar 
         Caption         =   "Bone Marrow Aspirate"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   285
         TabIndex        =   5
         Top             =   660
         Width           =   1470
      End
      Begin VB.CheckBox chk_fnac 
         Caption         =   "FNAC"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   3
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.ComboBox cmb_prev 
      Height          =   315
      Left            =   4980
      TabIndex        =   11
      Top             =   5145
      Width           =   3285
   End
   Begin VB.ComboBox cmb_later 
      Height          =   315
      ItemData        =   "frm_molecularreq.frx":0004
      Left            =   8550
      List            =   "frm_molecularreq.frx":000E
      TabIndex        =   2
      Top             =   3045
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.ComboBox cmb_site 
      Height          =   315
      Left            =   3120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   3030
      Width           =   4200
   End
   Begin VB.Shape Shape1 
      Height          =   4560
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   9375
   End
   Begin VB.Label Label7 
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
      Height          =   495
      Left            =   1560
      TabIndex        =   20
      Top             =   120
      Width           =   8700
   End
   Begin VB.Label Label6 
      Caption         =   "Prev Req No/Path No"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3165
      TabIndex        =   17
      Top             =   5175
      Width           =   1710
   End
   Begin VB.Label Label5 
      Caption         =   "Laterality"
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
      Left            =   7620
      TabIndex        =   16
      Top             =   3060
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label Label4 
      Caption         =   "Site/s"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2535
      TabIndex        =   15
      Top             =   3045
      Width           =   465
   End
   Begin VB.Label Label3 
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8400
      TabIndex        =   14
      Top             =   2280
      Width           =   405
   End
   Begin VB.Label Label2 
      Caption         =   "Procedure  "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1545
      TabIndex        =   13
      Top             =   2235
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Molecular Pathology Services"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   4320
      TabIndex        =   12
      Top             =   840
      Width           =   3690
   End
End
Attribute VB_Name = "molecularpathology"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim I1 As Integer
Dim SELRINVESTMP_QR As rdoQuery
Dim SELRINVESTMP_RS As rdoResultset
Dim SELRINVESTMPFLAG As Boolean
Dim SELMA101MP_QR As rdoQuery
Dim SELMA101MP_RS As rdoResultset
Dim SELMA101MP_FLG As Boolean
Dim selsite_qr As rdoQuery
Dim selsite_rs As rdoResultset
Dim selsite_flg As Boolean
Public ser_code As String

Private Sub cmd_back_Click()
 
 If chk_fnac.Value = 1 Then
    StrFresh = "FNAC"
 End If
 
 If chk_fblood.Value = 1 Then
   If StrFresh <> "" Then
    StrFresh = StrFresh & ", " & "Blood"
   Else
    StrFresh = "Blood"
   End If
 End If
 
 If chk_ftissue.Value = 1 Then
   If StrFresh <> "" Then
    StrFresh = StrFresh & ", " & "Tissue"
   Else
    StrFresh = "Tissue"
   End If
 End If
 
 If chk_bonemar.Value = 1 Then
   If StrFresh <> "" Then
    StrFresh = StrFresh & ", " & "Bone Marrow Aspiration"
   Else
    StrFresh = "Bone Marrow Aspiration"
   End If
 End If
 
 If chk_ablood.Value = 1 Then
     StrArchive = "Blood"
 End If
 
 If chk_atissue.Value = 1 Then
   If StrArchive <> "" Then
    StrArchive = StrArchive & ", " & "Tissue"
   Else
    StrArchive = "Tissue"
   End If
 End If
 
 If chk_paraffin.Value = 1 Then
   If StrArchive <> "" Then
    StrArchive = StrArchive & ", " & "Paraffin"
   Else
    StrArchive = "Paraffin"
   End If
 End If
 
 StrSite = Trim(cmb_site.Text)
 'StrLaterality = Trim(cmb_later.Text)
  
i = I1 - 1
While i >= 0
'S 'HW = list_code.List(I)
    If List_code.Selected(i) Then
        pathology.CMB_TEXT.AddItem List_code.List(i)
        List_code.Text = List_code.List(i)
        ser_code = List_code.List(i)
        ser_code1 = List_code.List(i) '23/02/10
    End If
    i = i - 1
Wend
'SER_CODE = pathology.CMB_TEXT.Text
  list_proc.clear
 'Me.Hide 'shweta 21-aug-09
 Unload Me
 pathology.bceltm.Visible = True
 pathology.cmb_cyprocedure.Visible = False
 pathology.cmb_cyprocode.Visible = False
 pathology.Label10.Visible = False
 pathology.Label11.Visible = False
End Sub

Private Sub Form_Load()

str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE=?"
Set SELMA101MP_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM RISLIB.RINVESTIGN WHERE MODALITYCD=?"
Set SELRINVESTMP_QR = rdb.CreateQuery("", str)


str = "select * from patoolslib.ma103"
Set selsite_qr = rdb.CreateQuery("", str)
selsite_flg = False

If SELRINVESTMPFLAG = False Then
        SELRINVESTMP_QR(0) = Trim(pathology.txt_modcode.Text)
        Set SELRINVESTMP_RS = SELRINVESTMP_QR.OpenResultset(1, 3)
        SELRINVESTMPFLAG = True
    Else
        SELRINVESTMP_QR(0) = Trim(pathology.txt_modcode.Text)
        SELRINVESTMP_RS.MoreResults
        SELRINVESTMP_RS.Requery
    End If
    If SELRINVESTMP_RS.RowCount > 0 Then
        list_proc.clear
        List_code.clear
        I1 = 0
        While Not SELRINVESTMP_RS.EOF
            List_code.AddItem Trim(SELRINVESTMP_RS("ser_code"))
            SELMA101MP_FLG = False
            If SELMA101MP_FLG = False Then
                SELMA101MP_QR(0) = Trim(SELRINVESTMP_RS("SER_CODE"))
                Set SELMA101MP_RS = SELMA101MP_QR.OpenResultset(1, 3)
                SELMA101MP_FLG = True
             Else
                SELMA101MP_QR(0) = SELRINVESTMP_RS("SER_CODE")
                SELMA101MP_RS.MoreResults
                SELMA101MP_RS.Requery
            End If
            
            If SELMA101MP_RS.RowCount > 0 Then
            
                While Not SELMA101MP_RS.EOF
                list_proc.AddItem SELMA101MP_RS("SER_DET")
                SELMA101MP_RS.MoveNext
                I1 = I1 + 1
                Wend
            End If
            SELRINVESTMP_RS.MoveNext
       Wend
    End If
    
     
  Set selsite_rs = selsite_qr.OpenResultset(2, 3)
  selsite_rs.MoveFirst
  While Not selsite_rs.EOF
       cmb_site.AddItem selsite_rs(1)
       selsite_rs.MoveNext
  Wend
    
End Sub
Private Sub list_proc_Click()
i = I1 - 1
While i >= 0
    If list_proc.Selected(i) Then
        List_code.Selected(i) = True
    ElseIf List_code.Selected(i) Then
        List_code.Selected(i) = False
    End If
    i = i - 1
Wend


End Sub

Private Sub List_code_Click()
i = I1 - 1
While i >= 0
    If List_code.Selected(i) Then
        list_proc.Selected(i) = True
    ElseIf list_proc.Selected(i) Then
        list_proc.Selected(i) = False
    End If
    i = i - 1
Wend

End Sub
