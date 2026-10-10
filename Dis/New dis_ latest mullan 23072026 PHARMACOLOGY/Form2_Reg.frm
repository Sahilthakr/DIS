VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9660
   LinkTopic       =   "Form2"
   ScaleHeight     =   8055
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   735
      Left            =   7800
      TabIndex        =   20
      Top             =   2280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&SAVE"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   19
      Top             =   7200
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Height          =   735
      Left            =   2040
      TabIndex        =   18
      Top             =   4920
      Width           =   5295
   End
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   2040
      TabIndex        =   17
      Top             =   3960
      Width           =   5295
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   2040
      TabIndex        =   16
      Top             =   3000
      Width           =   5295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   6240
      TabIndex        =   12
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   1440
      Width           =   5295
   End
   Begin VB.CommandButton cmd_exit 
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Top             =   7200
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox txt_time 
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
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   7920
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txt_date 
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
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   7920
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Kin's Details :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   15
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Permanent Address :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   14
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Address :"
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
      Left            =   360
      TabIndex        =   13
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Category :"
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
      Left            =   4440
      TabIndex        =   11
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Sex/Age :"
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
      Left            =   360
      TabIndex        =   8
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Name : "
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
      Left            =   360
      TabIndex        =   7
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "PATIENT REGISTRATION"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   840
      Width           =   4335
   End
   Begin VB.Label lbl_time 
      AutoSize        =   -1  'True
      Caption         =   "Time"
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
      Left            =   7320
      TabIndex        =   4
      Top             =   600
      Width           =   420
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Date"
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
      Left            =   7320
      TabIndex        =   3
      Top             =   240
      Width           =   540
   End
   Begin VB.Label LBL_TATA 
      Caption         =   "TATA MEMORIAL HOSPITAL"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim renv As rdoEnvironment
Dim rdb As rdoConnection
Dim dbcn As rdoConnection
Dim dbenv As rdoEnvironment
Dim rs_N As rdoResultset
Dim SEL_reg_QR As rdoQuery 'Added on 29/07/09 by Shraddha Pandit
Dim SEL_reg_RS As rdoResultset

Private Sub Command1_Click()
MaskEdBox1.S
End Sub

Private Sub Form_Load()
Main
Main1
datetime
txt_date.Text = Format(DATE11, "DD/MM/YYYY")
txt_time.Text = time1
'Added by Shraddha Pandit on 29/07/09
str = "select * from registration"
Set SEL_reg_QR = dbcn.CreateQuery("", str)
SEL_reg_QR.Execute
Set SEL_reg_RS = SEL_reg_QR.OpenResultset(1, 3)

While Not SEL_reg_RS.EOF
    MsgBox SEL_reg_RS("address")
    SEL_reg_RS.MoveNext
Wend

End Sub

Private Sub Main()

Dim conString As String
'conString = Nothing
'conn.Open ("dsn=regdsn;UID=sa;PWD=vid")
'conString = "Provider="sqloledb" ; user id=; password=; database=;"
'dim con as new ADODB.Connection(conString)
'con.open()
'Put a message box to display the message
'MsgBox ("ConnectionOpened")
'con.close()

'Connection with AS/400 DB2

Dim er As rdoError
On Error GoTo CnEh
Set renv = rdoEnvironments(0)
renv.CursorDriver = rdUseOdbc

Set rdb = renv.OpenConnection(DSNAME:="DISDSN", Prompt:=rdDriverNoPrompt, ReadOnly:=False, Connect:="System=10.1.4.2;UID=PATOOLS;PWD=PATOOLS;Database = EC_TMH")
LOGIN = "PATOOLS"

Exit Sub

CnEh:
    Debug.Print Err, Error
    For Each er In rdoErrors
        Debug.Print er.Description, er.Number
    Next er
    MsgBox "Either this user dosen't exist on Server or " & Chr(13) & " there is a communication link failure"
    MsgBox Error(Err)
End
End Sub
Private Sub Main1()
'Connection with SQL server 2005
'ESTABLISH THE CONNECTION
Dim er1 As rdoError
On Error GoTo cner
    Set dbenv = rdoEnvironments(0)
    dbenv.CursorDriver = rdUseOdbc

            Set dbcn = dbenv.OpenConnection(DSNAME:="REGDSN", Prompt:=rdDriverNoPrompt, ReadOnly:=False, Connect:="uid=sa;pwd=shraddha;database=TEST")
       Exit Sub
cner:
          MsgBox "failure!!!"
          End           ' established

End Sub

Private Sub datetime()
    Set rs_N = rdb.OpenResultset("SELECT DISTINCT CURRENT DATE DT1 , CURRENT TIME TM1 FROM PATOOLSLIB.MA107", 1, 3)
    DATE11 = rs_N("dt1")
    DATE11 = Format(DATE11, "DD/MM/YYYY")
    time1 = Format(rs_N("TM1"), "hh:mm:ss")
    DATETIMEFLAG = True
End Sub


Private Sub Timer1_Timer()
    datetime
    txt_time.Text = time1
End Sub

Private Sub CMD_EXIT_Click()
End
End Sub



