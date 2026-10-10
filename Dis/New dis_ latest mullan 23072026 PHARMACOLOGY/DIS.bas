Attribute VB_Name = "DIS"
Public req_sta As String
Public VAR As Variant
Public IHFLAG As Boolean
Public detail As String
Public StrFresh As String
Public StrArchive As String
Public StrSite As String
Public StrLaterality As String
Public StrCgdiag1 As String
Public StrCgdiag2 As String
Public FASPFLAG As Boolean
Public IPOPMEMO As Boolean
Public renv As rdoEnvironment
Public rdb As rdoConnection
Public RS  As rdoResultset
'Public CSF_FLAG As Boolean
Public ConInterface As rdoConnection
Public RSInterface As rdoResultset
Public DATE11 As String
Public time1 As String
Public str As String
Public SELMA10121_QR As rdoQuery
Public SELMA10121_RS As rdoResultset
Public SELMA10121FLAG As Boolean
Public pathologyFLAG As Boolean
Public DATETIMEFLAG As Boolean
Public SELMA106_QR As rdoQuery
Public SELMA106_RS As rdoResultset
Public SELMA106FLAG As Boolean
Public INNC As Boolean
Public IHTEST As Variant
Public usr_name As String
Public ipadd As String
Public SPREPORTFLAG As Boolean
Public IHREPORTFLAG As Boolean
Public ipmemoflag As Boolean
Public opmemoflag As Boolean
Public FAREPORTFLAG As Boolean
Public CNREPROTFLAG As Boolean
Public pathqflag As Boolean
Public CYREQFLAG As Boolean
Public CYREPORTFLAG As Boolean
Public MEMOFLAG As Boolean
Public FNACREQNO As String
Public SPFAREQNO As String
Public tmflag As Boolean
Public ELFLAG As Boolean
Public BCFLAG As Boolean
Public BCREPORTFLAG As Boolean
Public REQNO As String
Public meterial As String
Public bodyf As String
Public DDG As String
Public AM As String
Public ELREPORTFLAG As Boolean
Public TMREPORTFLAG As Boolean
'Public CSETPRINTER As New cSetDfltPrinter
Public MBREPORTFALG As Boolean
Public PATHQUERIESFLAG As Boolean
Public APPOINTMENTFLAG As Boolean
Public PFTESTFLAG As Boolean
Public HEMCBCFLAG As Boolean
Public HEMBMAFLAG As Boolean
Public WORKLIST As Boolean
Public PASSWORD As String 'used for biochemistry for checking password before modifying it.
' sundeep for pimpol  inserting into Req of BP

Public irrd As String

Public ser_code1 As String 'Added by Shraddha on 24/02/10

Public luco As String
Public IL As String
Public bpwb As String
Public bppc As String
Public bpsdp As String
Public bprdp As String
Public bpffp As String
Public bpfviiid As String
Public bpcryo As String
Public bpalb As String
Public bpremarks As String
' sundeep for pimpol  inserting into Req of BP
Public ind1 As String
Public ind2 As String
Public ind3 As String
Public ind4 As String
Public ind5 As String
Public ind6 As String
Public ind7 As String
Public ind8 As String
Public ind9 As String
Public tp1 As String
Public tp2 As String
Public tp3 As String
Public drug As String
Public dose As String
Public noofcycles As String
Public bsa As String
Public OTHERS As String
Public ECNEWFLAG As Boolean
Public EDVIEWFLAG As Boolean
Public ECMODFLAG As Boolean
Public QRYFLAG As Boolean
Public EXAMCOMPFLAG As Boolean
Public CHKFLAG As Boolean
Public frmendoscopyflag As Boolean
Public BFREPORTFLAG As Boolean
Public BMAREPORTFLAG As Boolean
Public HAEMATOLOGYFLAG As Boolean
Public HLFLAG As Boolean
Public BMFLAG As Boolean
Public BFFLAG As Boolean
Public HLAFLAG As Boolean
Public DTMFLAG As Boolean
Public BPFLAG As Boolean
Public DTMNO As String
Public PLTC As String
Public BG As String
Public BG1 As String
Public RFT As String
Public HB As String
Public aliquot As String 'Shraddha Pandit 06/03/09
Public wbunit As String 'Shraddha Pandit 06/03/09
Public pcunit As String 'Shraddha Pandit 06/03/09

Public QrySts As Boolean 'Shraddha Pandit 07/09/09

Public updFlfCD34 As Boolean 'Shraddha Pandit 22/09/09
Public insFlgCD34 As Boolean 'Shraddha Pandit 22/09/09

Public REQUION As String
Public TM As String
Public T_M As String
Public T_M1 As String
Public FLD As String
Public TYPEOFREQ As String
Public SR1 As String
Public SR2 As String
Public OTHERS1 As String
Public PHOBOTMIST As String
Public NOOFUNREP As String
''FOR HLATYPING
Public HTOther As String
Public DonorRelation As String
Public STRDonor As String
Public HLA_SEX As String
Public HLA_AGE As String

''FOR HLATYPING
'Public V_PROCE As Boolean

''*******SECURITY****************

Public UID As String
Public USRGRP As String
Public PASSWD As String
Public USRDEPT As String

Public USRNAME_QR As rdoQuery
Public USRNAME_RS As rdoResultset

Public LOGIN As String
''*******SECURITY****************

''****HAEMATOLOGY*********************

Public MPO As String
Public SBB As String
Public PAS As String
Public LAP As String
Public NSE As String
Public TRAP As String
Public cbc As String
Public hem As String
Public cg1 As String
Public cyto As String
Public smar As String
Public TT As String
Public oth As String
Public othtext As String
Public bfmode As String

Public PREVTAB As Integer

''****HAEMATOLOGY*********************
Public crossmatch As String

'V001 PATIENT DEPOSIT STARTS
Dim TA104_STSQRY As rdoQuery
Dim TA104_SRRS As rdoResultset
Public STAT As String
Public WT_AMT As Integer
Public RESUL As String

'''added by KANCHAN
Public showmain As Boolean
Public RNAGE1 As String
Public DATEREQ1 As Date
Public datereq As String
Public timereq As String

'******* Shraddha 25/05/2009 ******
Public HLA_MM As Boolean

Public CBCrec As Boolean 'Shraddha 23/11/09
Public panic As Boolean 'Shraddha 24/11/09
'**********************************
''' ADDED BY SANDHYA PROTOCOL START
Public MIRBNO As String
''' END OF PROTOCOL

'***** Shraddha 27/09/2011 *****
Public str_pbs As String
Public str_cytochem As String
Public str_surface As String
'********************************
'''' ACTREC & TMH REQUISITION & MEMO RECEIPTS.
'''' ADDED BY SANDHYA DATED 11.6.2012
'''' ADDED BY SANDHYA FOR SUSPENSE STATEMENT OF ACTREC & TMH
'Trust Receipts Details from TMH server

Public cCn As Boolean
Public renv1 As rdoEnvironment
Public ACTtmhCn As Boolean
Public rdbT As rdoConnection
Public SELACTMHREQ As rdoQuery
Public RSACTMHREQ As rdoResultset
Public INSACTTMHREQ As rdoQuery
Public INSACTMHMEMO As rdoQuery
Public UPDACTTMHREQ As rdoQuery
'''' ADDED BY SANDHYA DATED 18.7.2012 FOR BP RGJAY PATIENT
Public BPPATFLAG As Boolean
'-- Trust Amount
Public s_UpdTa113 As rdoQuery
'-- Insert trust deposit to TMH-(PATOOLSLIB/ACTRTEUST)
Public updACTTrt As rdoQuery
Public updTA113ta As rdoQuery
Public updTA113tb As rdoQuery
Public selACTTrt As rdoQuery
Public rsACTTrt As rdoResultset
Public updATRUSTa As rdoQuery
Public maxdateACT As Date
Dim selTmhMaxMemo As rdoQuery
Dim RsTmhMaxMemo As rdoResultset
Public selActMaxMemo As rdoQuery
Public RsActMaxMemo As rdoResultset
Public updTA115a As rdoQuery

Public selTA110c As rdoQuery
Public resTA110c As rdoResultset
Public selTA110cF As Boolean

Public MaxMemo As Boolean
Public maxdate As Date
Public selCompTMH1 As rdoQuery
Public rsCompTMH1 As rdoResultset
Public selCompTMH1F As Boolean
Public selCompTMH2 As rdoQuery
Public rsCompTMH2 As rdoResultset
Public selCompTMH2F As Boolean
Public selSumTCRMemo As rdoQuery
Public RsSumTCRMemo As rdoResultset
Public selCmpTMH As rdoQuery
Public RsCmpTMH As rdoResultset
Public selCmpCd As rdoQuery
Public RsCmpCd As rdoResultset
Public selTA107CR1 As rdoQuery
Public RsTA107CR1 As rdoResultset
Public selCRMemoT As rdoQuery
Public RsCRMemoT As rdoResultset
Public selFBSUMMARYCrT As rdoQuery
Public RsFBSUMMARYCrT As rdoResultset
Public DT0, TM0 As String

Dim depTOT1, depTOT2, depTOT3 As String
Dim serTOT1, serTOT2, selFrmDT, selToDT, trname, traccno, age1, sumMemo As String
Dim chkDupRecR, tmhRDO, DIS, chkDupS, Z100a, Z100t As Boolean
Public CRtOT, drTOT, blTOT As String




'V001 'PATIENT DEPOSIT ENDS

Public Function pad(l As Integer, ID As String, num As String, strpt As Integer) As String
Dim S
Dim ST1
Dim st2
Dim ST
Dim z
Dim z1
Dim c As Integer
Dim L1 As String
If "num" Like "R-######" Or "num" Like "M-######" Or "num" Like "B-######" Or "num" Like "F-######" Or "num" Like "I-######" Or "num" Like ">??/#####" Then
    MsgBox ("format correct")
Else
    L1 = Left(num, 2)
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
    'MsgBox num
    Select Case ID
    Case "R"
    st2 = "R-" & num
    Case "M"
    st2 = "M-" & num
    Case "C"
    st2 = L1 & "/" & num
    Case "B"
    st2 = "B-" & num
    Case "F"
    st2 = "F-" & num
    Case "I"
    st2 = "I-" & num
    End Select
    pad = st2
    'MsgBox pad
End Function
Public Function sendRMSG(ByVal REQNO As String, ByVal CASENO As String, ByVal reqstat As String, ByVal IPAD As String)
'If Trim(Left(CASENO, 1)) = "R" Or Trim(Left(CASENO, 1)) = "P" Then 'for rf and preventie cases due to smart card
'Else
Dim ob As smssend
Set ob = New sendsmsrpt.smssend
'ob.SENDMSG REQNO, CASENO, reqstat, IPAD 'varinder 2025
'End If
End Function

    Public Function req_chglen(str_reqno)

Dim res_reqno As String

 LEN1 = Len(Trim(str_reqno))

If LEN1 = 13 Then
res_reqno = Left(str_reqno, 3) & "/" & Mid(str_reqno, 4, 2) & "/" & Mid(str_reqno, 6, 2) & "/" & Right(str_reqno, 6)
 End If

If LEN1 = 16 Then
res_reqno = Left(str_reqno, 3) & Mid(str_reqno, 5, 2) & Mid(str_reqno, 8, 2) & Right(str_reqno, 6)
End If

If LEN1 = 27 Then
res_reqno = Left(str_reqno, 3) & Mid(str_reqno, 5, 2) & Mid(str_reqno, 8, 2) & Mid(str_reqno, 11, 6)
End If
req_chglen = res_reqno

End Function
Public Function GET_REVDATE(DATETEXT As String) As Variant
On Error GoTo MSG:
Dim K8, K81, KK8, KK881, KKK881, KK81
K8 = InStr(1, DATETEXT, "/")
If Not K8 = 0 Then
        K81 = InStr(K8, DATETEXT, "/")
        If Not K81 = 0 Then
        KK8 = (Mid(DATETEXT, 1, K8 - 1))
        KK8 = Trim(KK8)
        If Not KK8 = "" Then
            If CInt(KK8) = 0 Or CInt(KK8) > 31 Then
          '     MsgBox "INVALID DATE ENTERED i.e Days >31"
            GET_REVDATE = ""
                
                       Exit Function 'kk8=0 k8>31
            End If       'kk8=0 k8>31
           Else  'kk8=""
              'x'  MsgBox "INVALID DATE ENTERED i.e Month >12"
              GET_REVDATE = ""
                              
                Exit Function 'kk8=""
            End If       'kk8=""
        Else     'k81=0
     'x'   Exit Sub 'k81=0
        End If   'k81=0
        KK81 = InStr(K8 + 1, DATETEXT, "/")
        KK881 = (Mid(DATETEXT, K8 + 1, KK81 - 1 - K8))
        KK881 = Trim(KK881)
            If CInt(KK881) = 0 Or CInt(KK881) > 12 Then
             ' DATETEXT.Text = Format(DATE1, "dd/mm/yyyy")
            GET_REVDATE = ""
 'x'    MsgBox "INVALID DATE ENTERED i.e Month >12"

                Exit Function
            End If
    End If

KKK881 = (Mid(DATETEXT, KK81 + 1))
GET_REVDATE = CStr(KKK881 + "-" + KK881 + "-" + KK8)

MSG:

End Function
Public Sub ACTREQMAIN()
Dim er As rdoError
On Error GoTo CnEh
Dim DT01 As String
Dim TM01 As String
ACTtmhCn = False

Set renv1 = rdoEnvironments(0)
renv1.CursorDriver = rdUseOdbc
Dim VARBADRI01 As String

Set rdbT = renv1.OpenConnection(DSNAME:="ACTRECTESTS", Prompt:=rdDriverNoPrompt, ReadOnly:=False, Connect:="UID=TMH;PWD=TMH;Database = PATOOLSLIB")
Call ACTTMHQRY
'Call actreceiptqry
ACTtmhCn = True
Exit Sub

CnEh:
    Debug.Print Err, ERROR
    For Each er In rdoErrors
        Debug.Print er.Description, er.Number
    Next er
    MsgBox "Either this user dosen't exist on Server or " & Chr(13) & " there is a TMH AS/400 server communication link failure"
    MsgBox ERROR(Err)
    tmhCn = False
    End
End Sub
Public Sub ACTTMHQRY()
'''' ACTREC & TMH REQUISITION & MEMO RECEIPTS.
'''' ADDED BY SANDHYA DATED 1.6.2012
Dim str As String
str = "SELECT * FROM PATOOLSLIB.ACTTMHREQ WHERE MEMONO = ? "
Set SELACTMHREQ = rdbT.CreateQuery("", str)
SELACTMHREQ(0) = "0000000"
Set RSACTMHREQ = SELACTMHREQ.OpenResultset(1, 1)

str = "insert into patoolslib.acttmhmemo(memo_no,memo_date,ser_code,ser_amt,qty,transtat,ipaddrs,userid,recstat) values (?,?,?,?,?,?,?,?,?)"
Set INSACTMHMEMO = rdbT.CreateQuery("insactmem", str)

str = "INSERT INTO PATOOLSLIB.ACTTMHREQ(CASE_NO,REQNNO,REQNDT,DEPT_CODE,MODALITYCD,PATNSTAT,PAT_CATG,IPADDR,USERID,TRDATE,TRTIME) VALUES(?,?,?,?,?,?,?,?,?,?,?)"
Set INSACTTMHREQ = rdbT.CreateQuery("", str)

str = "update PATOOLSLIB.ACTTMHREQ SET REQNSTAT=?,RECPTNO=?,RECPTDT=?,MEMONO=?,MEMODT=?,MEMO_AMT=? where REQNNO=?"
Set UPDACTTMHREQ = rdbT.CreateQuery("", str)

str = "UPDATE PATOOLSLIB.TA113 SET SER_AMT = SER_AMT + ? WHERE (CASE_NO = ?)"
Set updTA113t = rdbT.CreateQuery("updTA13t", str)

str = "SELECT * FROM PATOOLSLIB.TA113 WHERE CASE_NO = ?"
Set selTA113c = rdbT.CreateQuery("selTA113c", str)

str = "SELECT ack_DATE AS RECDT FROM PATOOLSLIB.TA1131 WHERE ACK_NO = ? ORDER BY ACK_DATE "
Set selTA1131c = rdbT.CreateQuery("selTA1131c", str)
selTA1131c(0) = "00099999"
Set rsTA1131c = selTA1131c.OpenResultset(1, 1)


End Sub
Public Sub ACTTMH_TABal(CASENO As String)
' Padmakar - (TMH-ACTREC) joint company account details
'Trust Balance Check from TMH-ACTREC Account

Dim s_ACTRTRt As rdoQuery
Dim r_ACTRTRt As rdoResultset
Dim f_ACTRTRt As Boolean

Dim s_TA110t As rdoQuery
Dim r_TA110t As rdoResultset
Dim f_TA110t As Boolean

Dim s_TACRBANKt As rdoQuery
Dim r_TACRBANKt As rdoResultset
Dim f_TACRBANKt As Boolean

Dim s_FBSUMMARYt As rdoQuery
Dim r_FBSUMMARYt As rdoResultset
Dim f_FBSUMMARYt As Boolean

Dim s_FBSUMMARYa As rdoQuery
Dim r_FBSUMMARYa As rdoResultset
Dim f_FBSUMMARYa As Boolean

Dim s_TA107t As rdoQuery
Dim r_TA107t As rdoResultset
Dim f_TA107t As Boolean

Dim s_TA107a As rdoQuery
Dim r_TA107a As rdoResultset
Dim f_TA107a As Boolean
    
Dim s_FOXBALt As rdoQuery
Dim r_FOXBALt As rdoResultset
Dim f_FOXBALt As Boolean

' Collect Foxpro balance amount from FOXBAL as a total deposit in patient's trust account
Set s_FOXBALt = rdb.CreateQuery("s_FOXBALt", "SELECT * FROM PATOOLSLIB.FOXBAL WHERE CASE_NO = ?")
s_FOXBALt(0) = "TT/10001"
Set r_FOXBALt = s_FOXBALt.OpenResultset(2, 3)

s_FOXBALt(0) = CASENO
r_FOXBALt.MoreResults
r_FOXBALt.Requery
depTOT1 = 0
If r_FOXBALt.RowCount > 0 Then
depTOT1 = depTOT1 + (Val(r_FOXBALt("NET_AMT")) - Val(r_FOXBALt("SER_AMT")))
End If

' Collect 677/1207/2860 and 438/1206 receipt sum from ACTRTRUST as a total deposit in patient's trust account
Set s_ACTRTRt = rdb.CreateQuery("s_ACTRTRt", "SELECT IFNULL(SUM(REC_AMT), 0) AS recTOT FROM PATOOLSLIB.ACTRTRUST WHERE CASE_NO = ? AND TRANSTAT <> 'Y' AND PRN_FLAG = 'P'")
s_ACTRTRt(0) = "TT/10001"
Set r_ACTRTRt = s_ACTRTRt.OpenResultset(2, 3)

s_ACTRTRt(0) = CASENO
r_ACTRTRt.MoreResults
r_ACTRTRt.Requery
'depTOT1 = 0
If r_ACTRTRt.RowCount > 0 Then
depTOT1 = depTOT1 + Val(r_ACTRTRt("recTOT"))
End If

' Collect 677/1207 and 438/1206 receipt sum from TA110 as a total deposit in patient's trust account, which is not in ACTRTRUST (Old Receipts)
Set s_TA110t = rdb.CreateQuery("s_TA11t", "SELECT IFNULL(SUM(REC_AMT), 0) AS recTOT FROM PATOOLSLIB.TA110 A, PATOOLSLIB.TA1101 B WHERE CASE_NO = ? AND (A.REC_NO = B.REC_NO) AND a.prn_flag = 'P' AND (B.ACC_NUM = 677 or (B.ACC_NUM = 438 and dep_flag='T')) AND A.REC_NO NOT IN (SELECT REC_NO FROM PATOOLSLIB.ACTRTRUST)")
s_TA110t(0) = "TT/10001"
Set r_TA110t = s_TA110t.OpenResultset(2, 3)

s_TA110t(0) = CASENO
r_TA110t.MoreResults
r_TA110t.Requery
If r_TA110t.RowCount > 0 Then
depTOT1 = depTOT1 + Val(r_TA110t("recTOT"))
End If


str = "SELECT IFNULL(SUM(REC_AMT), 0) AS recTOT FROM PATOOLSLIB.TACRBANK WHERE CASE_NO = ? "
Set s_TACRBANKt = rdb.CreateQuery("s_TACRBANKt", str)
s_TACRBANKt(0) = "TT/10001"
Set r_TACRBANKt = s_TACRBANKt.OpenResultset(2, 3)

s_TACRBANKt(0) = CASENO
r_TACRBANKt.MoreResults
r_TACRBANKt.Requery
If r_TACRBANKt.RowCount > 0 Then
depTOT1 = depTOT1 + Val(r_TACRBANKt("recTOT"))
End If

'''' Collect 677/1207 and 438/1206 receipt sum from ACTRTRUST as a total deposit in patient's trust account
'''STR = "SELECT IFNULL(SUM(REC_AMT), 0) AS recTOT FROM PATOOLSLIB.ACTRTRUST WHERE CASE_NO = ? and PRN_FLAG = 'P' AND (ACC_NUM = 1207 OR ACC_NUM = 677 OR ACC_NUM = 1206 OR ACC_NUM = 438) AND ARECDATE BETWEEN ? AND ?"
'''Set sACTRTR = rdb.CreateQuery("sACTRTR", STR)
'''
'''sACTRTR(0) = CASENO
'''sACTRTR(1) = DT0
'''sACTRTR(2) = DT0
'''Set rACTRTR = sACTRTR.OpenResultset(1, 1)
'''depTOT1 = rACTRTR("recTOT")


'calculate total bill amount all Trust Bills, Interim Bills and Final Bill for Turst Patient
'Sum Trust Bill amount from TMH

Call selFrmToDt(CASENO)

str = "SELECT IFNULL(SUM(NET_AMT), 0) AS fbAMT FROM PATOOLSLIB.FBSUMMARY WHERE CASE_NO = ? and PAT_TYPE = 'TA' AND FB_DATE BETWEEN ? AND ?"
Set s_FBSUMMARYt = rdbT.CreateQuery("s_FBSUMMARYt", str)

depTOT2 = 0

s_FBSUMMARYt(0) = CASENO
s_FBSUMMARYt(1) = selFrmDT
s_FBSUMMARYt(2) = selToDT
Set r_FBSUMMARYt = s_FBSUMMARYt.OpenResultset(1, 1)
If r_FBSUMMARYt.RowCount > 0 Then
depTOT2 = depTOT2 + Val(r_FBSUMMARYt("fbAMT"))
End If

str = "SELECT IFNULL(SUM(NET_AMT), 0) AS fbAMT FROM PATOOLSLIB.FBSUMMARY WHERE CASE_NO = ? and PAT_TYPE = 'TA' AND FB_DATE BETWEEN ? AND ?"
Set s_FBSUMMARYa = rdb.CreateQuery("s_FBSUMMARYa", str)

'Sum Trust Bill amount from ACTREC
s_FBSUMMARYa(0) = CASENO
s_FBSUMMARYa(1) = selFrmDT
s_FBSUMMARYa(2) = selToDT
Set r_FBSUMMARYa = s_FBSUMMARYa.OpenResultset(1, 1)
If r_FBSUMMARYa.RowCount > 0 Then
depTOT2 = depTOT2 + Val(r_FBSUMMARYa("fbAMT"))
End If


' -- START -- Total not billed pending memos collected here from TMH and ACTREC
' Sum all Not Billed (pending) Memos for Trust Patient
    
str = "SELECT IFNULL(SUM(MEMO_AMT), 0) AS memoAMT FROM PATOOLSLIB.TA107 WHERE CASE_NO = ? and AMT_RCVD = 'T' AND (MEMO_DATE BETWEEN ? AND ?) AND MEMO_NO NOT IN (SELECT DISTINCT(MEMO_NO) FROM PATOOLSLIB.TA1071 WHERE (ser_code = 'Z006') )" '
Set s_TA107t = rdbT.CreateQuery("s_TA107t", str)

' Sum Memo from TMH
depTOT3 = 0

s_TA107t(0) = CASENO
s_TA107t(1) = selFrmDT
s_TA107t(2) = selToDT
Set r_TA107t = s_TA107t.OpenResultset(1, 1)
If r_TA107t.RowCount > 0 Then
depTOT3 = depTOT3 + Val(r_TA107t("memoAMT"))
End If

str = "SELECT IFNULL(SUM(MEMO_AMT), 0) AS memoAMT FROM PATOOLSLIB.TA107 WHERE CASE_NO = ? and AMT_RCVD = 'T' AND (MEMO_DATE BETWEEN ? AND ?) AND MEMO_NO NOT IN (SELECT DISTINCT(MEMO_NO) FROM PATOOLSLIB.TA1071 WHERE (ser_code = 'Z006') )"
Set s_TA107a = rdb.CreateQuery("s_TA107a", str)
' Sum Memo from ACTREC
s_TA107a(0) = CASENO
s_TA107a(1) = selFrmDT
s_TA107a(2) = selToDT
Set r_TA107a = s_TA107a.OpenResultset(1, 1)
If r_TA107a.RowCount > 0 Then
depTOT3 = depTOT3 + Val(r_TA107a("memoAMT"))
End If

' -- END -- Total not billed pending memos collected here from TMH and ACTREC
' Calculate total Credit amount, total Debit amount and total Balance amount
CRtOT = depTOT1
drTOT = depTOT2 + depTOT3
blTOT = CRtOT - drTOT
End Sub
Public Sub chkLastMemo(S1 As String)
' Padmakar - (TMH-ACTREC) joint company account details
str = "SELECT MAX(MEMO_DATE) FROM PATOOLSLIB.TA107 WHERE CASE_NO=?"
Set selTmhMaxMemo = rdbT.CreateQuery("selmhMaxMemo", str)

selTmhMaxMemo(0) = Trim(S1)
Set RsTmhMaxMemo = selTmhMaxMemo.OpenResultset(1, 1)
    If IsNull(RsTmhMaxMemo(0)) = False Then
       MaxMemo = True
       maxdateACT = RsTmhMaxMemo(0)
    Else
       MaxMemo = False
       maxdateACT = "01/01/1900"
    End If
End Sub
Private Function selFrmToDt(S1 As String) As String
' Padmakar - (TMH-ACTREC) joint company account details
selTA110c(0) = Trim(S1)
If selTA110cF Then
   resTA110c.MoreResults
   resTA110c.Requery
   selTA110cF = False
Else
Set resTA110c = selTA110c.OpenResultset(1, 1)
    selTA110cF = True
End If
If Not resTA110c.EOF Then
  If IsNull(resTA110c("RECDT")) = True Then
    selFrmDT = "01" & Mid(DATE11, 3, 10)
    selToDT = DATE11
  Else
   selFrmDT = resTA110c("RECDT")
   selToDT = date1
  End If
Else
   selFrmDT = ""
   selToDT = ""
End If
End Function

Public Function tmhDRQuery()
Dim str As String
' Padmakar - (TMH-ACTREC) joint company account details
'Queries for TMH DR
str = "SELECT * FROM PATOOLSLIB.TA115"
Set selCompTMH1 = rdbT.CreateQuery("selCompTMH1", str)

str = "SELECT * FROM PATOOLSLIB.MCOMPANY"
Set selCompTMH2 = rdbT.CreateQuery("selCompTMH2", str)

'STR = "SELECT MAX(MEMO_DATE) FROM PATOOLSLIB.TA107 WHERE CASE_NO=?"
'Set selTmhMaxMemo = rdbT.CreateQuery("selTmhMaxMemo", STR)

'STR = "SELECT sum(MEMO_amt) FROM PATOOLSLIB.TA107 WHERE CASE_NO = ? and memo_date > ?"
'Set selSumTCRMemo = rdbT.CreateQuery("selSumTCRMemo", STR)

str = "SELECT * FROM PATOOLSLIB.TA107 WHERE CASE_NO = ? and memo_date > ?"
Set selSumTCRMemo = rdbT.CreateQuery("selSumTCRMemo", str)

str = "SELECT * FROM PATOOLSLIB.TA114 WHERE CASE_NO = ?"
Set selCmpTMH = rdbT.CreateQuery("selCmpTMH", str)

str = "SELECT * FROM PATOOLSLIB.TA115 WHERE COMP_CODE = ?"
Set selCmpCd = rdbT.CreateQuery("selCmpCd", str)

str = "SELECT * FROM PATOOLSLIB.TA107 WHERE CASE_NO = ? and AMT_RCVD = 'C' AND (MEMO_DATE BETWEEN ? AND ?) AND MEMO_NO NOT IN (SELECT DISTINCT(MEMO_NO) FROM PATOOLSLIB.TA1071 WHERE (ser_code = 'Z006') )" '
Set selTA107CR1 = rdbT.CreateQuery("selTA107CR1", str)

str = "SELECT * FROM PATOOLSLIB.TA1071 WHERE MEMO_NO = ?"
Set selCRMemoT = rdbT.CreateQuery("selCRMemoT", str)

' All Company Bills for Case No
str = "SELECT * FROM PATOOLSLIB.FBSUMMARY WHERE CASE_NO = ? and PAT_TYPE = 'CR' AND FB_DATE BETWEEN ? AND ?"
Set selFBSUMMARYCrT = rdbT.CreateQuery("selFBSUMMARYCrT", str)

End Function

Public Function char_valid(ST As Integer) As Integer
Char = Chr(ST)
ST = Asc(UCase(Char))
char_valid = ST
End Function
Public Sub Main()

Dim er As rdoError
'On Error GoTo CnEh
Set renv = rdoEnvironments(0)
renv.CursorDriver = rdUseOdbc
       'Set rdb = renv.OpenConnection(dsname:="DISDSN", Prompt:=rdDriverNoPrompt, ReadOnly:=False, Connect:="UID=PATOOLS;PWD=PATOOLS;Database = EC_TMH")
       'Set rdb = renv.OpenConnection(dsname:="Disdsn", Prompt:=rdDriverNoPrompt, ReadOnly:=False, Connect:="UID=" & UID & ";PWD=" & UID & ";Database = EC_TMH") 'THIS CODE IS USED TO CONTECT TO WITH SECURITY.
       ' Set rdb = renv.OpenConnection(dsname:="DISDSN", Prompt:=rdDriverNoPrompt, ReadOnly:=False, Connect:="UID=" & Trim(nlogonfrm.OLDTEXT) & ";PWD=" & Trim(nlogonfrm.OLDTEXT) & ";Database = EC_TMH") 'THIS CODE IS USED TO CONTECT TO WITH SECURITY.
        
                
                
        If Mid(Trim(UID), 2) <> "DISCLK" Then
            
            '*********** hardcoded ip address as 10.1.4.2 on date 03.08.2006
            'Set rdb = renv.OpenConnection(DSNAME:="DISDSN", Prompt:=rdDriverNoPrompt, ReadOnly:=False, Connect:="System=10.1.4.2;UID=" & Trim(UID) & ";PWD=" & Trim(UID) & ";Database = EC_TMH") 'THIS CODE IS USED TO CONTECT TO WITH SECURITY.
            'Commented and addded by Shraddha on 26/10/10
            Set rdb = renv.OpenConnection(DSNAME:="DISDSN", Prompt:=rdDriverNoPrompt, ReadOnly:=False, Connect:="UID=" & Trim(UID) & ";PWD=" & Trim(UID) & ";Database = DISLIB") 'THIS CODE IS USED TO CONTECT TO WITH SECURITY.
            'Set rdb = renv.OpenConnection(dsname:="DISDSN", Prompt:=rdDriverNoPrompt, ReadOnly:=False, Connect:="UID=" & Trim(UID) & ";PWD=" & Trim(UID) & ";Database = EC_TMH") 'THIS CODE IS USED TO CONTECT TO WITH SECURITY.
            LOGIN = Trim(UID)
        Else
            Set rdb = renv.OpenConnection(DSNAME:="DISDSN", Prompt:=rdDriverNoPrompt, ReadOnly:=False, Connect:="System=10.1.4.2;UID=PATOOLS;PWD=PATOOLS;Database = EC_TMH")
            LOGIN = "PATOOLS"
        End If
Exit Sub

CnEh:
         Debug.Print Err, ERROR
        For Each er In rdoErrors
            Debug.Print er.Description, er.Number
        Next er
        MsgBox "Either this user dosen't exist on Server or " & Chr(13) & " there is a communication link failure"
    MsgBox ERROR(Err)
    End
    

End Sub
Public Sub DATETIME()
    Set RS = rdb.OpenResultset("SELECT DISTINCT CURRENT DATE DT1 , CURRENT TIME TM1 FROM PATOOLSLIB.MA107") '***2/2/99
    DATE11 = RS("dt1") '**add 24/11/2k
   DATE11 = Format(DATE11, "DD/MM/YYYY")
       time1 = Format(RS("TM1"), "hh:mm:ss")
    DATETIMEFLAG = True '***2/2/99
    

    'DATE11 = Date
End Sub

Public Sub selta101()
'STR = "select * from PATOOLSLIB.ta101 where case_no=?"
'Set selta101_qr = rdb.CreateQuery("selta101_qr1", STR)

End Sub

Public Sub SELMA106()
    str = "SELECT * FROM PATOOLSLIB.MA106"
    Set SELMA106_QR = rdb.CreateQuery("SELMA106_QR1", str)

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
    'MsgBox num
   
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
    'MsgBox pad
 
End Function


Public Sub CreateDSN()
Dim EN As rdoEnvironment
Dim CNTEST As rdoConnection
Dim str As String
str = "Description=" _
        & "Client Access Express ODBC data source" _
  & Chr$(13) & "OemToAnsi=No" _
  & Chr$(13) & "SERVER=TMH810" _
  & Chr$(13) & "Network=DBNMPNTW" _
  & Chr$(13) & "Database=EC_TMH" _
  
' Create new registered DSN.
rdoEngine.rdoRegisterDataSource "Example", "TMH810", True, str
' Open the database.

End Sub

'V001 PATIENT DEPOSIT STARTS
Public Function WAIT_STAT(STAT As String, MCNO As String, WT_AMT As Long)
Set TA104_STSQRY = rdb.CreateQuery("TA104_STSQRYSCA", "SELECT * FROM PATOOLSLIB.TA104 WHERE CASE_NO = ? AND WT_TYPE = 'D' ORDER BY BED_TYPE ")
TA104_STSQRY(0) = MCNO
Set TA104_SRRS = TA104_STSQRY.OpenResultset(1, 1)
'1B
 If TA104_SRRS.RowCount = 0 Then
     STAT = "1"
     WT_AMT = 0
'1EL
   Else
     STAT = "0"
      '2B
          If Trim(TA104_SRRS("BED_TYPE")) = "D" Then
              WT_AMT = 5000
             ElseIf Trim(TA104_SRRS("BED_TYPE")) = "P" Then
              WT_AMT = 2500
            ElseIf Trim(TA104_SRRS("BED_TYPE")) = "SP" Then
              WT_AMT = 1000
          End If
      '2E
 End If
'1E
TA104_SRRS.Close
TA104_STSQRY.Close
End Function
'''' Added by Sandhya for New numbering system '''' dated 8.9.2022
'''' Added new function NEWNOPAD for paddding zeros in caseno.
Public Function NEWNOpad(ByVal CASENO As String) As String
Dim K As Integer
        Dim str1 As String
        Dim str As String
        Dim padding As String
        If InStr(CASENO, "/") = 0 Then
            CASENO = Trim(CASENO)
        Else
            CASENO = Replace(CASENO, "/", "")
        End If
        K = Len(Trim(CASENO))
        padding = ""

        If IsNumeric(Left(CASENO, 1)) Then
            If K = 8 Then
                str1 = Mid(CASENO, 1, 7)
                str = Mid(CASENO, 8, K)
                padding = UCase(str1) + "/00000" + str
            ElseIf K = 9 Then
                str1 = Mid(CASENO, 1, 7)
                str = Mid(CASENO, 8, K)
                padding = UCase(str1) + "/0000" + str
            ElseIf K = 10 Then
                str1 = Mid(CASENO, 1, 7)
                str = Mid(CASENO, 8, K)
                padding = UCase(str1) + "/000" + str
            ElseIf K = 11 Then
                str1 = Mid(CASENO, 1, 7)
                str = Mid(CASENO, 8, K)
                padding = UCase(str1) + "/00" + str
            ElseIf K = 12 Then
                str1 = Mid(CASENO, 1, 7)
                str = Mid(CASENO, 8, K)
                padding = UCase(str1) + "/0" + str
            ElseIf K = 13 Then
                str1 = Mid(CASENO, 1, 7)
                str = Mid(CASENO, 8, K)
                padding = UCase(str1) + "/" + str
            End If
            NEWNOpad = Trim(padding)
        Else
            If K = 3 Then
                str1 = Mid(CASENO, 1, 2)
                str = Mid(CASENO, 3, K)
                padding = UCase(str1) + "/0000" + str
            ElseIf K = 4 Then
                str1 = Mid(CASENO, 1, 2)
                str = Mid(CASENO, 3, K)
                padding = UCase(str1) + "/000" + str
            ElseIf K = 5 Then
                str1 = Mid(CASENO, 1, 2)
                str = Mid(CASENO, 3, K)
                padding = UCase(str1) + "/00" + str
            ElseIf K = 6 Then
                str1 = Mid(CASENO, 1, 2)
                str = Mid(CASENO, 3, K)
                padding = UCase(str1) + "/0" + str
            ElseIf K = 7 Then
                'str1 = Left(caseno, 2)
                'str = Mid(caseno, 4, 5)
                str1 = Mid(CASENO, 1, 2)
                str = Mid(CASENO, 3, K)
                padding = UCase(str1) + "/" + str
              '  padding = UCase(CASENO)
            End If
            NEWNOpad = Trim(padding)
        End If
End Function

Public Sub insertIntoRtimeLog(REQNO As String)
    Dim rsq As rdoQuery
    STRSQL = "Insert into rtimelog"
    Set rsq = rdb.CreateQuery("InsRtimeLog", STRSQL)
End Sub
'V001 PATIENT DEPOSIT ENDS
