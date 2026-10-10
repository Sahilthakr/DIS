Attribute VB_Name = "SMSSEND"
Dim renv As rdoEnvironment
Dim rdb As rdoConnection
Dim RS  As rdoResultset

Dim SELSMSTBL_QR As rdoQuery
Dim SELSMSTBL_RS As rdoResultset

Dim SELMOBNO_QR As rdoQuery
Dim SELMOBNO_RS As rdoResultset



Public Sub Main1()



Set renv = rdoEnvironments(0)
renv.CursorDriver = rdUseOdbc

Set rdb = renv.OpenConnection(dsname:="DISDSN", Prompt:=rdDriverNoPrompt, ReadOnly:=False, Connect:="UID=TMH;PWD=TMH;Database = EC_TMH")

Exit Sub


End Sub



Sub main()

'Main1
 'SENDMSG "FZZSP13000123", "IT/00001"
    
End Sub


Sub SENDMSG(ByVal REQNNO As String, ByVal CASENO As String, ByVal reqnstat As String, ByVal IPAD As String)
Main1

Dim SELEMRMODALIT_QR As rdoQuery
Dim SELEMRMODALIT_RS As rdoResultset
REQNO = REQNNO
Dim rs_server As rdoResultset
Dim rs_smsdtls As rdoResultset
Dim DEPT_NM As String
Dim txtmsg As String
Dim dbsql As New ADODB.Connection
Dim str As String
Dim msgid1 As String
If reqnstat = "F" Then
msgid1 = "16"
ElseIf reqnstat = "T" Then
msgid1 = "17"
ElseIf reqnstat = "A" Then
msgid1 = "18"
Else
Exit Sub
End If
 SelStr = "SELECT * FROM PATOOLSLIB.TBLSMSMSGS WHERE MSGID='" & Trim(msgid1) & "'"
 Set SELSMSTBL_QR = rdb.CreateQuery("", SelStr)
  Set SELSMSTBL_RS = SELSMSTBL_QR.OpenResultset(1, 1)
    If SELSMSTBL_RS.EOF = False Then
    
            SelStr = "SELECT * FROM PATOOLSLIB.EMRMODALIT WHERE DEPT_CODE='" & Trim(Left(REQNNO, 3)) & "' AND MODALITYCD='" & Trim(Mid(REQNNO, 4, 2)) & "'"
        Set SELEMRMODALIT_QR = rdb.CreateQuery("", SelStr)
        Set SELEMRMODALIT_RS = SELEMRMODALIT_QR.OpenResultset(1, 1)
        DEPT_NM = ""
        If SELEMRMODALIT_RS.EOF = False Then
        DEPT_NM = Trim(SELEMRMODALIT_RS("MODLITYDES"))
        End If
    
    
        If Trim(SELSMSTBL_RS("MSGDESC")) <> "" And Len(Trim(SELSMSTBL_RS("MSGDESC"))) <> 0 Then
           MSG = Trim(SELSMSTBL_RS("MSGDESC"))
          '  MSG = Trim(SELSMSTBL_RS("MSGDESC"))
            
             
            
                MSG = Replace(MSG, "@1", Trim(DEPT_NM))
                MSG = Replace(MSG, "@2", Trim(REQNNO))
                MSG = Replace(MSG, "@3", Trim(CASENO))
            
            Set rs_smsdtls = rdb.OpenResultset("SELECT * FROM PATOOLSLIB.VWSMSCASES WHERE CASENO='" & CASENO & "' ", 1, 1)
            If Not rs_smsdtls.EOF Then
                If Trim(rs_smsdtls("pmob")) <> "" Then
                   
                     Set rs_server = rdb.OpenResultset("select server from patoolslib.serverip where APPLICATN ='SMS'", 1, 1)
                        
                      '  dbsql.Open ("Provider=SQLOLEDB.1;Password=tmhit123;Persist Security Info=True;User ID=SA;Initial Catalog=SMS_DB;Data Source=" & Trim(rs_server(0)))
                        'Set rq_ins = dbsql.CreateQuery("", "insert into SMS_IO_LOG_MASTER(SMS_MOBILE,SMS_TEXT,STAMP,PDU_SMS,PRIORITY_INDEX,SMS_METHOD,SMS_IO_IND) values ('" & Trim(rs_smsdtls("MOBILE1")) & "','" & Trim(msg) & "',getdate(),0,'" & Trim(rs_smsdtls("priority")) & "','2','3')")
                       ' str = "insert into SMS_IO_LOG_MASTER(SMS_MOBILE,SMS_TEXT,STAMP,PDU_SMS,PRIORITY_INDEX,SMS_METHOD,SMS_IO_IND,transaction_id,case_no,msgID) values ('" & Trim(rs_smsdtls("pmob")) & "','" & Trim(MSG) & "',getdate(),0,'" & Trim(SELSMSTBL_RS("priority")) & "','2','3','" & REQNNO & "','" & CASENO & "','" & Trim(msgid1) & "')"
                                         
                        
                       ' dbsql.Execute str
                       ' dbsql.Close
                        
                        str = "insert into DISLIB.RPTSMSLOG(REQNO,CASE_NO,REQSTAT,SMSDEL_DT,SMSDEL_TM,FLAG,IPADD,MSGID) values ('" & Trim(REQNNO) & "','" & Trim(CASENO) & "','" & Trim(reqnstat) & "',CURRENT DATE,CURRENT TIME,'Y','" & Trim(IPAD) & "','" & Trim(msgid1) & "')"
                       MsgBox (str)
                        rdb.Execute str
                 End If
            End If
         End If
    End If
End Sub


