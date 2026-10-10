VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form Smpl_Acc 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Sample Accession Record"
   ClientHeight    =   9450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12645
   LinkTopic       =   "Form2"
   ScaleHeight     =   9450
   ScaleWidth      =   12645
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.OptionButton opt_other 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Other Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5640
      TabIndex        =   19
      Top             =   3480
      Width           =   2655
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   6975
      Left            =   1320
      TabIndex        =   2
      Top             =   1200
      Width           =   8775
      Begin VB.Frame Frame2 
         BackColor       =   &H00808080&
         Caption         =   "Select Report : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         Left            =   1080
         TabIndex        =   6
         Top             =   600
         Width           =   6975
         Begin VB.CommandButton CMDREPORT 
            BackColor       =   &H0088B7B7&
            Caption         =   "Show &Report"
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
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   3840
            Width           =   1815
         End
         Begin VB.CommandButton CMD_CANCEL 
            BackColor       =   &H0088B7B7&
            Caption         =   "&Cancel"
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
            Left            =   1080
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   3840
            Width           =   1335
         End
         Begin VB.CommandButton CMD_BACK 
            BackColor       =   &H0088B7B7&
            Caption         =   "&Back"
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
            Left            =   4920
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   3840
            Width           =   1335
         End
         Begin VB.OptionButton OPT_5 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Positive Results"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   600
            TabIndex        =   11
            Top             =   1680
            Width           =   2415
         End
         Begin VB.OptionButton OPT_4 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Routine Micro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3240
            TabIndex        =   10
            Top             =   1200
            Width           =   1935
         End
         Begin VB.OptionButton Opt_3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Total Test Count"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   600
            TabIndex        =   9
            Top             =   1200
            Width           =   2055
         End
         Begin VB.OptionButton Opt_2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Result List Report"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   3240
            TabIndex        =   8
            Top             =   600
            Width           =   2175
         End
         Begin VB.OptionButton Opt_1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Sample Accession Record"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   600
            Left            =   600
            TabIndex        =   7
            Top             =   480
            Width           =   2415
         End
         Begin MSMask.MaskEdBox msk_fromdt 
            Height          =   375
            Left            =   2280
            TabIndex        =   14
            Top             =   3000
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   8959927
            ForeColor       =   4210752
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox msk_todt 
            Height          =   375
            Left            =   4680
            TabIndex        =   15
            Top             =   3000
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   8959927
            ForeColor       =   0
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label4 
            BackColor       =   &H80000018&
            Caption         =   "From Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   255
            Left            =   1200
            TabIndex        =   18
            Top             =   3000
            Width           =   1215
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H80000018&
            Caption         =   "To Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   195
            Left            =   3840
            TabIndex        =   16
            Top             =   3000
            Width           =   705
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000018&
            BorderColor     =   &H00000000&
            BorderStyle     =   6  'Inside Solid
            Height          =   2175
            Left            =   480
            Shape           =   4  'Rounded Rectangle
            Top             =   2520
            Width           =   6015
         End
      End
      Begin VB.TextBox TXT_MESSAGE 
         Appearance      =   0  'Flat
         BackColor       =   &H0088B7B7&
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
         Height          =   360
         Left            =   2280
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   5880
         Width           =   5655
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   1560
         TabIndex        =   3
         Top             =   3000
         Width           =   6255
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00808080&
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         Height          =   6735
         Left            =   -120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   8895
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "MESSAGE"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   5880
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport CR 
      Left            =   240
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   840
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lbl_req 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "DEPARTMENT OF MICROBIOLOGY : SECTION SEROLOGY"
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
      Height          =   330
      Left            =   1920
      TabIndex        =   1
      Top             =   720
      Width           =   8205
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
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
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   11055
   End
End
Attribute VB_Name = "Smpl_Acc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

''''' Designed and Coded by Kanchan
Dim RS As rdoResultset
Dim str As String
Dim flag1 As Boolean
Dim TESTLISTFLAG As Boolean

'shweta 06-oct-2012
Dim SELTA101_QR As rdoQuery
Dim SELTA101_RS As rdoResultset


Dim SELDMULPROCEAC_QR As rdoQuery
Dim SELDMULPROCEAC_RS As rdoResultset

Dim SELDREQUISITNAC_QR As rdoQuery
Dim SELDREQUISITNAC_RS As rdoResultset

Dim SELPROCODE_QR As rdoQuery
Dim SELPROCODE_RS As rdoResultset

Dim INSTMPACCESS_QR As rdoQuery
Dim INSTMPACCESS_RS As rdoResultset

Dim INSTESTLIST_QR As rdoQuery
Dim DELTSTELIST_QR As rdoQuery

Dim SELREQNO_QR As rdoQuery
Dim SELREQNO_RS As rdoResultset
Dim SELLABNO_QR As rdoQuery
Dim SELLABNO_RS As rdoResultset

Dim DELTMPACCESS_QR As rdoQuery
Dim SELRREQ_QR As rdoQuery
Dim SELRREQ_RS As rdoResultset
Dim INSMICRO_TEMP As rdoQuery
Dim SELMICROPROCE_QR As rdoQuery
Dim SELMICROPROCE_RS As rdoResultset
Dim DEL_MICROTEMP_QR As rdoQuery
Dim a As Integer
Dim EN As HMSCrypt


Private Sub cmd_back_Click()
    
    msk_fromdt.Mask = ""
    msk_fromdt.Text = ""
    msk_fromdt.Mask = "##/##/####"
    msk_todt.Mask = ""
    msk_todt.Text = ""
    msk_todt.Mask = "##/##/####"
     
    DIS_MAIN.show
    Me.Hide
   
    CMDREPORT.Enabled = True

End Sub

Private Sub cmd_cancel_Click()
msk_fromdt.Text = "__/__/____"
msk_todt.Text = "__/__/____"
msk_fromdt.Text = Date
msk_todt.Text = Date

'Msk_caseno.Mask = "??>/#####"
'Msk_caseno.Text = ""
'CMD_CANCEL.Enabled = False

OPT_1.Value = False
OPT_2.Value = False
OPT_3.Value = False

CMDREPORT.Enabled = True
End Sub


Private Sub CMDREPORT_Click()

 Dim LEN1 As Integer
            Dim LEN2 As Integer
            Dim LEN3 As Integer
            Dim l As Integer
            Dim str1 As String


   TXT_MESSAGE.Text = "  DISPLAYING REPORT PLEASE WAIT......"
    MYDELAY (5)
   
'''''''''     Sample Accession Record
If OPT_1.Value = True Then
    Dim var_date_f As String
    Dim var_date_t As String
    Dim RS_STNAME As rdoResultset
    
    var_date_f = Year(msk_fromdt.FormattedText) & "," & month(msk_fromdt.FormattedText) & "," & Day(msk_fromdt.FormattedText)
    var_date_t = Year(msk_todt.FormattedText) & "," & month(msk_todt.FormattedText) & "," & Day(msk_todt.FormattedText)


    If flag1 = False Then
        SELDMULPROCEAC_QR(0) = var_date_f
        SELDMULPROCEAC_QR(1) = var_date_t
    
        Set SELDMULPROCEAC_RS = SELDMULPROCEAC_QR.OpenResultset(1, 3)
        flag1 = True
    Else
        SELDMULPROCEAC_QR(0) = var_date_f
        SELDMULPROCEAC_QR(1) = var_date_t
    
        SELDMULPROCEAC_RS.MoreResults
        SELDMULPROCEAC_RS.Requery
    End If

        While Not SELDMULPROCEAC_RS.EOF
    
            INSTMPACCESS_QR(1) = ""
            INSTMPACCESS_QR(2) = ""
            INSTMPACCESS_QR(3) = ""
            
            INSTMPACCESS_QR(0) = SELDMULPROCEAC_RS(0)
        
              Dim req As String
                req = SELDMULPROCEAC_RS(0)
        
            '' TO SELECT PROCODE
                SELPROCODE_QR(0) = req
                Set SELPROCODE_RS = SELPROCODE_QR.OpenResultset(1, 3)
            
            
            INSTMPACCESS_QR(1) = "No"
            INSTMPACCESS_QR(2) = "No"
            INSTMPACCESS_QR(3) = "No"
            INSTMPACCESS_QR(4) = "No"
            INSTMPACCESS_QR(5) = "No"
            INSTMPACCESS_QR(6) = "No"
            INSTMPACCESS_QR(7) = "No"
            INSTMPACCESS_QR(8) = "No"
            INSTMPACCESS_QR(9) = "No"
            INSTMPACCESS_QR(10) = "No"
            INSTMPACCESS_QR(11) = "No"
            INSTMPACCESS_QR(12) = "No"
            INSTMPACCESS_QR(13) = "No"
            INSTMPACCESS_QR(14) = "No"
            INSTMPACCESS_QR(15) = "No"
            INSTMPACCESS_QR(18) = "No"
                While Not SELPROCODE_RS.EOF
                                     
                    'G171
                        If SELPROCODE_RS(0) = "G171" Then
                            INSTMPACCESS_QR(1) = "Yes"
                        End If
                        
                    'G129
                        If SELPROCODE_RS(0) = "G129" Then
                            INSTMPACCESS_QR(2) = "Yes"
                        End If
                    
                    'G134
                        If SELPROCODE_RS(0) = "G134" Then
                            INSTMPACCESS_QR(3) = "Yes"
                        End If
                        
                        If SELPROCODE_RS(0) = "G126" Then
                            INSTMPACCESS_QR(4) = "Yes"
                        End If
                        
                        If SELPROCODE_RS(0) = "G133" Then
                            INSTMPACCESS_QR(5) = "Yes"
                        End If
                        
                        If SELPROCODE_RS(0) = "G130" Then
                            INSTMPACCESS_QR(6) = "Yes"
                        End If
                        
                        If SELPROCODE_RS(0) = "G136" Then
                            INSTMPACCESS_QR(7) = "Yes"
                        End If
                        
                        If SELPROCODE_RS(0) = "G254" Then
                            INSTMPACCESS_QR(8) = "Yes"
                        End If
                        
                        If SELPROCODE_RS(0) = "G255" Then
                            INSTMPACCESS_QR(9) = "Yes"
                        End If
                        
                        If SELPROCODE_RS(0) = "G132" Then
                            INSTMPACCESS_QR(10) = "Yes"
                        End If
                        
                        If SELPROCODE_RS(0) = "G131" Then
                            INSTMPACCESS_QR(11) = "Yes"
                        End If
                        
                        If SELPROCODE_RS(0) = "G121" Then
                            INSTMPACCESS_QR(12) = "Yes"
                        End If
                        
                        If SELPROCODE_RS(0) = "G122" Then
                            INSTMPACCESS_QR(13) = "Yes"
                        End If
                        
                        If SELPROCODE_RS(0) = "G161" Then
                            INSTMPACCESS_QR(14) = "Yes"
                        End If
                        
                        If SELPROCODE_RS(0) = "G162" Then
                            INSTMPACCESS_QR(15) = "Yes"
                        End If
                        
                        If SELPROCODE_RS(0) = "G261" Then
                            INSTMPACCESS_QR(18) = "Yes"
                        End If
                    SELPROCODE_RS.MoveNext
                Wend
        
            INSTMPACCESS_QR(16) = SELDMULPROCEAC_RS(1)
        
        ''TO SELECT PATHNO
        
            SELDREQUISITNAC_QR(0) = req
            Set SELDREQUISITNAC_RS = SELDREQUISITNAC_QR.OpenResultset(1, 3)
        
            
           
            str1 = ""
            LEN1 = Len(Trim(SELDREQUISITNAC_RS(0)))
            LEN2 = 13 - LEN1
            
            For l = 1 To LEN2
            
                str1 = str1 & "0"
                
            Next
            
            str1 = str1 & SELDREQUISITNAC_RS(0)
            
            LEN3 = Len(Trim(str1))
            INSTMPACCESS_QR(17) = Trim(str1)
        
            SELTA101_QR(0) = Trim(SELDREQUISITNAC_RS(1))
            Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
            
            
            If SELTA101_RS.EOF = False Then
            INSTMPACCESS_QR(19) = Trim(SELTA101_RS("TAG")) & " " & Trim(SELTA101_RS("FIRST_NAME")) & " " & Trim(SELTA101_RS("LAST_NAME"))
            Else
            Set RS_STNAME = rdb.OpenResultset("SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO='" & Trim(SELDMULPROCEAC_RS("REQNNO")) & "'")
                If RS_STNAME.EOF = False Then
                INSTMPACCESS_QR(19) = Trim(RS_STNAME("PNAME"))
                Else
                INSTMPACCESS_QR(19) = ""
                End If
            End If
            
            INSTMPACCESS_QR.Execute
            SELDMULPROCEAC_RS.MoveNext
        Wend
    
    
''''''
    CR.ReportFileName = "C:\DISREPORTSEC\SampleAccess.rpt"
    CR.Formulas(0) = "from=Date(" & var_date_f & ")"
    CR.Formulas(1) = "to=Date(" & var_date_t & ")"
    
    'CR.SelectionFormula = "{DMULPROCE.REQNDT}>=Date(" & var_date_f & ") AND {DMULPROCE.REQNDT}<=Date(" & var_date_t & ")"
    CR.WindowState = crptMaximized
    CR.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
    CR.Action = 1
      
        DELTMPACCESS_QR.Execute
  
End If


'***** Shraddha 07/12/11 *****
If opt_other.Value = True Then
    'Dim var_date_f As String
    'Dim var_date_t As String

    var_date_f = Year(msk_fromdt.FormattedText) & "," & month(msk_fromdt.FormattedText) & "," & Day(msk_fromdt.FormattedText)
    var_date_t = Year(msk_todt.FormattedText) & "," & month(msk_todt.FormattedText) & "," & Day(msk_todt.FormattedText)


    'If flag1 = False Then
        SELDMULPROCEAC_QR(0) = var_date_f
        SELDMULPROCEAC_QR(1) = var_date_t
    
        Set SELDMULPROCEAC_RS = SELDMULPROCEAC_QR.OpenResultset(1, 3)
     '   flag1 = True
    'Else
     '   SELDMULPROCEAC_QR(0) = var_date_f
      '  SELDMULPROCEAC_QR(1) = var_date_t
    
       ' SELDMULPROCEAC_R1S.MoreResults
        'SELDMULPROCEAC_RS.Requery
    'End If

        While Not SELDMULPROCEAC_RS.EOF
    
            INSTMPACCESS_QR(1) = ""
            INSTMPACCESS_QR(2) = ""
            INSTMPACCESS_QR(3) = ""
            
            INSTMPACCESS_QR(0) = SELDMULPROCEAC_RS(0)
        
              'Dim req As String
                req = SELDMULPROCEAC_RS(0)
        
            '' TO SELECT PROCODE
            
                SELPROCODE_QR(0) = req
                Set SELPROCODE_RS = SELPROCODE_QR.OpenResultset(1, 3)
            If req = "GZZSE11033853" Then
                MsgBox "Hiii"
            End If
            
            INSTMPACCESS_QR(1) = "No"
            INSTMPACCESS_QR(2) = "No"
            INSTMPACCESS_QR(3) = "No"
            INSTMPACCESS_QR(4) = "No"
            INSTMPACCESS_QR(5) = "No"
            INSTMPACCESS_QR(6) = "No"
            INSTMPACCESS_QR(7) = "No"
            INSTMPACCESS_QR(8) = "No"
            INSTMPACCESS_QR(9) = "No"
            INSTMPACCESS_QR(10) = "No"
            INSTMPACCESS_QR(11) = "No"
            INSTMPACCESS_QR(12) = "No"
            INSTMPACCESS_QR(13) = "No"
            INSTMPACCESS_QR(14) = "No"
            INSTMPACCESS_QR(15) = "No"
            INSTMPACCESS_QR(18) = "No"
                While Not SELPROCODE_RS.EOF
                                     
                    'G171
                        If SELPROCODE_RS(0) = "G171" Then
                            INSTMPACCESS_QR(1) = "Yes"
'                        Else
'                            INSTMPACCESS_QR(1) = "No"
                        End If
                        
                    'G129
                        If SELPROCODE_RS(0) = "G129" Then
                            INSTMPACCESS_QR(2) = "Yes"
'                        Else
'                            INSTMPACCESS_QR(2) = "No"
                        End If
                    
                    'G134
                        If SELPROCODE_RS(0) = "G134" Then
                            INSTMPACCESS_QR(3) = "Yes"
'                        Else
'                            INSTMPACCESS_QR(3) = "No"
                        End If
                        
                        If SELPROCODE_RS(0) = "G126" Then
                            INSTMPACCESS_QR(4) = "Yes"
'                        Else
'                            INSTMPACCESS_QR(4) = "No"
                        End If
                        
                        If SELPROCODE_RS(0) = "G133" Then
                            INSTMPACCESS_QR(5) = "Yes"
'                        Else
'                            INSTMPACCESS_QR(5) = "No"
                        End If
                        
                        If SELPROCODE_RS(0) = "G130" Then
                            INSTMPACCESS_QR(6) = "Yes"
'                        Else
'                            INSTMPACCESS_QR(6) = "No"
                        End If
                        
                        If SELPROCODE_RS(0) = "G136" Then
                            INSTMPACCESS_QR(7) = "Yes"
'                        Else
'                            INSTMPACCESS_QR(7) = "No"
                        End If
                        
                        If SELPROCODE_RS(0) = "G254" Then
                            INSTMPACCESS_QR(8) = "Yes"
'                        Else
'                            INSTMPACCESS_QR(8) = "No"
                        End If
                        
                        If SELPROCODE_RS(0) = "G255" Then
                            INSTMPACCESS_QR(9) = "Yes"
'                        Else
'                            INSTMPACCESS_QR(9) = "No"
                        End If
                        
                        If SELPROCODE_RS(0) = "G132" Then
                            INSTMPACCESS_QR(10) = "Yes"
'                        Else
'                            INSTMPACCESS_QR(10) = "No"
                        End If
                        
                        If SELPROCODE_RS(0) = "G131" Then
                            INSTMPACCESS_QR(11) = "Yes"
'                        Else
'                            INSTMPACCESS_QR(11) = "No"
                        End If
                        
                        If SELPROCODE_RS(0) = "G121" Then
                            INSTMPACCESS_QR(12) = "Yes"
'                        Else
'                            INSTMPACCESS_QR(12) = "No"
                        End If
                        
                        If SELPROCODE_RS(0) = "G122" Then
                            INSTMPACCESS_QR(13) = "Yes"
'                        Else
'                            INSTMPACCESS_QR(13) = "No"
                        End If
                        
                        If SELPROCODE_RS(0) = "G161" Then
                            INSTMPACCESS_QR(14) = "Yes"
'                        Else
'                            INSTMPACCESS_QR(14) = "No"
                        End If
                        
                        If SELPROCODE_RS(0) = "G162" Then
                            INSTMPACCESS_QR(15) = "Yes"
                            End If
                        If SELPROCODE_RS(0) = "G261" Then
                            INSTMPACCESS_QR(16) = "Yes"
'                        Else
'                            INSTMPACCESS_QR(15) = "No"
                        End If
                    SELPROCODE_RS.MoveNext
                Wend
        
            INSTMPACCESS_QR(16) = SELDMULPROCEAC_RS(1)
        
        ''TO SELECT PATHNO
        
            SELDREQUISITNAC_QR(0) = req
            Set SELDREQUISITNAC_RS = SELDREQUISITNAC_QR.OpenResultset(1, 3)
        
            
           
            str1 = ""
            LEN1 = Len(Trim(SELDREQUISITNAC_RS(0)))
            LEN2 = 13 - LEN1
            
            For l = 1 To LEN2
            
                str1 = str1 & "0"
                
            Next
            
            str1 = str1 & SELDREQUISITNAC_RS(0)
            
            LEN3 = Len(Trim(str1))
            INSTMPACCESS_QR(17) = Trim(str1)
        
        
            INSTMPACCESS_QR.Execute
            SELDMULPROCEAC_RS.MoveNext
        Wend
    
    
''''''
    CR.ReportFileName = "C:\DISREPORTSEC\SampleAccessOtherReport.rpt"
    CR.Formulas(0) = "from=Date(" & var_date_f & ")"
    CR.Formulas(1) = "to=Date(" & var_date_t & ")"
    CR.WindowState = crptMaximized
    CR.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
    CR.Action = 1
      
        DELTMPACCESS_QR.Execute
  
End If
'********* Shraddha **********

''''''''''      Result List Records
If OPT_2.Value = True Then

    Dim var_date_f1 As String
    Dim var_date_t1 As String

    var_date_f1 = Year(msk_fromdt.FormattedText) & "," & month(msk_fromdt.FormattedText) & "," & Day(msk_fromdt.FormattedText)
    var_date_t1 = Year(msk_todt.FormattedText) & "," & month(msk_todt.FormattedText) & "," & Day(msk_todt.FormattedText)

    If TESTLISTFLAG = False Then
        SELREQNO_QR(0) = var_date_f1
        SELREQNO_QR(1) = var_date_t1
    
        Set SELREQNO_RS = SELREQNO_QR.OpenResultset(1, 3)
        TESTLISTFLAG = True
    Else
        SELREQNO_QR(0) = var_date_f1
        SELREQNO_QR(1) = var_date_t1
        
        SELREQNO_RS.MoreResults
        SELREQNO_RS.Requery
    End If
    
    Dim CNT As Integer
    
    CNT = SELREQNO_RS.RowCount
    
    While Not SELREQNO_RS.EOF
    
            INSTESTLIST_QR(0) = SELREQNO_RS(0)
            
            SELDREQUISITNAC_QR(0) = SELREQNO_RS(0)
            Set SELDREQUISITNAC_RS = SELDREQUISITNAC_QR.OpenResultset(1, 3)
        
                      
                    str1 = ""
                    LEN1 = Len(Trim(SELDREQUISITNAC_RS(0)))
                    LEN2 = 7 - LEN1
            
                    For l = 1 To LEN2
            
                        str1 = str1 & "0"
                
                    Next
            
                    str1 = str1 & SELDREQUISITNAC_RS(0)
            
                    LEN3 = Len(Trim(str1))
        
            INSTESTLIST_QR(1) = Trim(str1)
            INSTESTLIST_QR.Execute
        
        SELREQNO_RS.MoveNext
    Wend
    
    CR1.ReportFileName = "C:\DISREPORTSEC\TESTRESULT.rpt"
    CR1.Formulas(0) = "from=Date(" & var_date_f1 & ")"
    CR1.Formulas(1) = "to=Date(" & var_date_t1 & ")"
    CR1.SelectionFormula = "{MINTERFACE.TDATE}>=Date(" & var_date_f1 & ") AND {MINTERFACE.TDATE}<=Date(" & var_date_t1 & ")"
    CR1.WindowState = crptMaximized
    CR1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
    CR1.Action = 1
    
    DELTSTELIST_QR.Execute

End If

''''''''           Total List Count
If OPT_3.Value = True Then

    Dim var_date_f2 As String
    Dim var_date_t2 As String

    var_date_f2 = Year(msk_fromdt.FormattedText) & "," & month(msk_fromdt.FormattedText) & "," & Day(msk_fromdt.FormattedText)
    var_date_t2 = Year(msk_todt.FormattedText) & "," & month(msk_todt.FormattedText) & "," & Day(msk_todt.FormattedText)

    CR1.ReportFileName = "C:\DISREPORTSEC\TESTCOUNT.rpt"
    CR1.Formulas(0) = "from=Date(" & var_date_f2 & ")"
    CR1.Formulas(1) = "to=Date(" & var_date_t2 & ")"
    CR1.SelectionFormula = "{MINTERFACE.TDATE}>=Date(" & var_date_f2 & ") AND {MINTERFACE.TDATE}<=Date(" & var_date_t2 & ")"
    CR1.WindowState = crptMaximized
    CR1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
    CR1.Action = 1

    CR1.ReportFileName = "C:\DISREPORTSEC\TESTCNT.rpt"
    CR1.Formulas(0) = "from=Date(" & var_date_f2 & ")"
    CR1.Formulas(1) = "to=Date(" & var_date_t2 & ")"
    
    CR1.SelectionFormula = "{MINTERFACE.TDATE}>=Date(" & var_date_f2 & ") AND {MINTERFACE.TDATE}<=Date(" & var_date_t2 & ")"
    CR1.WindowState = crptMaximized
    CR1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
    CR1.Action = 1

End If

If OPT_4.Value = True Then

    Dim var_date_f3 As String
    Dim var_date_t3 As String
    
    var_date_f3 = Year(msk_fromdt.FormattedText) & "," & month(msk_fromdt.FormattedText) & "," & Day(msk_fromdt.FormattedText)
    var_date_t3 = Year(msk_todt.FormattedText) & "," & month(msk_todt.FormattedText) & "," & Day(msk_todt.FormattedText)
    
    CR1.ReportFileName = "C:\disreportsec\RoutineMicRO1.rpt"
    CR1.Formulas(0) = "from=Date(" & var_date_f3 & ")"
    CR1.Formulas(1) = "to=Date(" & var_date_t3 & ")"
    CR1.SelectionFormula = "{RTIMELOG.SAMRECDATE}>=DATE(" & var_date_f3 & ") AND {RTIMELOG.SAMRECDATE}<=DATE(" & var_date_t3 & ") AND LEFT(TRIM({RTIMELOG.REQNNO}),5)='GZZBA'"
    CR1.WindowState = crptMaximized
    CR1.Connect = "ODBC;DATABASE=DISLIB;UID=" & Trim(LOGIN) & ";PWD=" & Trim(LOGIN) & ";DSN=DISDSN"
    CR1.Action = 1
End If

If OPT_5.Value = True Then

    Dim var_date_f4 As String
    Dim var_date_t4 As String
    
    var_date_f4 = Year(msk_fromdt.FormattedText) & "," & month(msk_fromdt.FormattedText) & "," & Day(msk_fromdt.FormattedText)
    var_date_t4 = Year(msk_todt.FormattedText) & "," & month(msk_todt.FormattedText) & "," & Day(msk_todt.FormattedText)
    
    SELRREQ_QR(0) = msk_fromdt.FormattedText
    SELRREQ_QR(1) = msk_todt.FormattedText
    
    Set SELRREQ_RS = SELRREQ_QR.OpenResultset(1, 3)
          
    If SELRREQ_RS.RowCount > 0 Then
    
        While Not SELRREQ_RS.EOF
        
            SELMICROPROCE_QR(0) = SELRREQ_RS("REQNNO")
            Set SELMICROPROCE_RS = SELMICROPROCE_QR.OpenResultset(1, 3)
        Dim ST1 As String
            If SELMICROPROCE_RS.RowCount > 0 Then
                While Not SELMICROPROCE_RS.EOF
                    a = CInt(SELMICROPROCE_RS("FLD1"))
                    ST1 = EN.Decrypt(Trim(SELMICROPROCE_RS("RESULT")), a)
                    If UCase(EN.Decrypt(Trim(SELMICROPROCE_RS("RESULT")), a)) = "REACTIVE" Or UCase(EN.Decrypt(Trim(SELMICROPROCE_RS("RESULT")), a)) = "HIGH REACTIVE" Then
                    
                        INSMICRO_TEMP(0) = SELRREQ_RS("REQNNO")
                        INSMICRO_TEMP(1) = SELMICROPROCE_RS("PRO_CODE")
                        INSMICRO_TEMP(2) = UCase(EN.Decrypt(Trim(SELMICROPROCE_RS("OBSERVALUE")), a))
                        INSMICRO_TEMP(3) = UCase(EN.Decrypt(Trim(SELMICROPROCE_RS("RESULT")), a))
                        
                        INSMICRO_TEMP.Execute
                    End If
                    
                    SELMICROPROCE_RS.MoveNext
                Wend
            End If
        
            SELRREQ_RS.MoveNext
        Wend
    
    CR1.ReportFileName = "C:\disreportsec\MICROPOSITIVE.rpt"
    CR1.Formulas(0) = "from=Date(" & var_date_f4 & ")"
    CR1.Formulas(1) = "to=Date(" & var_date_t4 & ")"
    CR1.WindowState = crptMaximized
    CR1.Connect = "ODBC;DATABASE=DISLIB;UID=" & Trim(LOGIN) & ";PWD=" & Trim(LOGIN) & ";DSN=DISDSN"
    CR1.Action = 1
    
    DEL_MICROTEMP_QR.Execute
    End If
    
End If

   TXT_MESSAGE.Text = "  Click on BACK button to EXIT......"
    MYDELAY (5)


End Sub

Private Sub Form_Activate()
 msk_fromdt.Text = Date
 msk_todt.Text = Date
 TXT_MESSAGE.Text = ""
 'MSK_FROMDT1.Text = Date
 'MSK_TODT1.Text = Date
End Sub

Private Sub Form_Load()
 
Set RS = rdb.OpenResultset("SELECT DISTINCT CURRENT DATE DT1 , CURRENT TIME TM1 FROM PATOOLSLIB.MA107") '***2/2/99
DATE11 = RS("dt1")
 
 
'STR = "SELECT DISTINCT REQNO,REQNDT FROM DISLIB.DMULPROCE WHERE REQNDT >= ? AND REQNDT <= ? AND REQNO LIKE 'GZZSE%'"
'Set SELDMULPROCEAC_QR = rdb.CreateQuery("SELDMULPROCE", STR)

'str = "SELECT DISTINCT REQNO,TDATE, FROM INTERFACE.MINTERFACE WHERE TDATE >= ? AND TDATE <= ? AND REQNO LIKE 'GZZSE%'"
str = "SELECT DISTINCT REQNNO,SAMRECDATE FROM RISLIB.RTIMELOG WHERE SAMRECDATE >= ? AND SAMRECDATE <= ? AND REQNNO LIKE 'GZZSE%'"
Set SELDMULPROCEAC_QR = rdb.CreateQuery("SELDMULPROCE", str)


str = "SELECT DISTINCT REQNO FROM INTERFACE.MINTERFACE WHERE TDATE >= ? AND TDATE <= ? AND REQNO LIKE 'GZZSE%'"
Set SELREQNO_QR = rdb.CreateQuery("SELECTREQNUM", str)

str = "INSERT INTO DISLIB.TESTLIST (REQNO,LABNO) VALUES (?,?)"
Set INSTESTLIST_QR = rdb.CreateQuery("inserttestlist", str)

str = "DELETE FROM DISLIB.TESTLIST"
Set DELTSTELIST_QR = rdb.CreateQuery("DELETETESTLIST", str)

str = "SELECT PATHNO,CASENO FROM DISLIB.DREQUISITN WHERE REQNO=?"
Set SELDREQUISITNAC_QR = rdb.CreateQuery("SELDREQUISITN", str)
 
'STR = "SELECT PRO_CODE FROM DISLIB.DMULPROCE WHERE REQNO=?"
'Set SELPROCODE_QR = rdb.CreateQuery("SELPROCODE", STR)
 
str = "SELECT TESTCODE FROM INTERFACE.MINTERFACE WHERE REQNO=?"
Set SELPROCODE_QR = rdb.CreateQuery("SELPROCODE", str)
 
'str = "INSERT INTO DISLIB.TMPACCESS (REQNO,G171,G129,G134,REQNDT,LABNO) VALUES(?,?,?,?,?,?)"
'Commented and added by Shraddha on 07/12/11
str = "INSERT INTO DISLIB.TMPACCESS (REQNO,G171,G129,G134,G126,G133,G130,G136,G254,G255,G132,G131,G121,G122,G161,G162,REQNDT,LABNO,G261,name) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
Set INSTMPACCESS_QR = rdb.CreateQuery("INSERTQR", str)

str = "DELETE FROM DISLIB.TMPACCESS"
Set DELTMPACCESS_QR = rdb.CreateQuery("DELTETMP", str)

str = "SELECT *  FROM PATOOLSLIB.TA101 WHERE CASE_NO=?"
Set SELTA101_QR = rdb.CreateQuery("SELTA101", str)



'''FOR OPT_5
Set EN = New HMSCrypt

str = "SELECT * FROM RISLIB.RREQUISITN WHERE (FREPORTDT>=? AND FREPORTDT<=?) AND REQNNO LIKE 'GZZSE%'"
Set SELRREQ_QR = rdb.CreateQuery("SELRREQ", str)

str = "INSERT INTO DISLIB.MICR_TEMP1 (REQNO,PRO_CODE,OBSERVALUE,RESULT) VALUES (?,?,?,?)"
Set INSMICRO_TEMP = rdb.CreateQuery("INSERTMICTMP", str)


str = "SELECT * FROM DISLIB.MICROPROCE WHERE REQNO=?"
Set SELMICROPROCE_QR = rdb.CreateQuery("SELMICPROCE", str)

str = "DELETE FROM DISLIB.MICR_TEMP1"
Set DEL_MICROTEMP_QR = rdb.CreateQuery("DELMICROTEMP", str)
 
 msk_fromdt.Text = Date
 msk_todt.Text = Date
  
OPT_1.Value = False
OPT_2.Value = False
OPT_3.Value = False

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




