Dim  Input
Dim strRecset, strLine, strDescr, strSQL, strComment, strInfo, strdistinguishedName, strB1, strB2, strZK, strZK2, strmsproxyAddresses 
Dim arrDescr, arrComment, arrInfo, arrproxyAddresses 

Const ADS_SCOPE_SUBTREE = 2

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE

Set db = CreateObject("ADODB.Connection")
db.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=IAM;Data Source=GMZRZSQC070"
IF db.State = 1 Then
	db.execute("truncate Table [IAM].[dbo].[t_AD_User_ZusatzDaten]")

	Set objRecordSetHelp = db.Execute("SELECT [B1],[B2],[strZK],[strZK2] FROM [IAM].[dbo].[t_AD_User_help]" )
	objRecordSetHelp.MoveFirst
	Do Until objRecordSetHelp.EOF
		strB1 = objRecordSetHelp.Fields("B1").Value
		strB2 = objRecordSetHelp.Fields("B2").Value
		strZK = objRecordSetHelp.Fields("strZK").Value
		strZK2 = objRecordSetHelp.Fields("strZK2").Value

		objCommand.CommandText = "SELECT distinguishedName, description, comment, info, proxyAddresses FROM 'LDAP://gmzrzdc001.ferchau.local:3268' WHERE objectCategory='person' and objectClass = 'user' and samaccountname " & StrZK & " '" & strB1 & "' and samaccountname " & strZK2 & " '" & strB2 & "'" 
			Set objRecordSet = objCommand.Execute
		objRecordSet.MoveFirst

		Do Until objRecordSet.EOF
			strdistinguishedName =  Replace(objRecordSet.Fields("distinguishedName").Value,"'","`")
			arrDescr = objRecordSet.Fields("description").Value
			If IsNull(arrDescr) Then
				strDescr = "NULL"
			Else
				For Each strLine In arrDescr
					strDescr = "'" & Replace(strLine,"'","`") & "'"
				Next
			End If
			arrComment = objRecordSet.Fields("comment").Value
			If IsNull(arrComment) Then
				strComment = "NULL"
			Else
				strComment = "'" & Replace(strLine,"'","`") & "'"
			End If
			arrInfo = objRecordSet.Fields("info").Value
			If IsNull(arrInfo) Then
				strInfo = "NULL"
			Else
				strInfo = "'" & Replace(strLine,"'","`") & "'"
			End If
			arrInfo = objRecordSet.Fields("proxyAddresses").Value
			If IsNull(arrmsproxyAddresses) Then
				strmsproxyAddresses = "NULL"
			Else
				strmsproxyAddresses = "'" & Replace(strmsproxyAddresses,"'","`") & "'"
			End If

			strSQL = "insert into [IAM].[dbo].[t_AD_User_ZusatzDaten] (description,comment,info,proxyAddresses,distinguishedName) values (" & strDescr & "," & strComment & "," & strInfo & "," & strmsproxyAddresses & ",'" & strdistinguishedName & "')"
			db.Execute(strSQL)
			objRecordSet.MoveNext
		Loop
		objRecordSetHelp.MoveNext
	LOOP
Else
	IF Err.Number <> 0 Then
		db.Execute("insert into dbo.t_logfile(Meldungstext) values('Error # " & CStr(Err.Number) & " - " & Err.Description & "')")
		Err.Clear
	End If
End If	
db.close
objConnection.close
