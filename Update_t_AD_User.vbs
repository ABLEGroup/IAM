Dim  Input
Dim strRecset, strLine, strDescr, strSQL, strComment, strInfo, strdistinguishedName
Dim arrDescr, arrComment, arrInfo

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
	objCommand.CommandText = "SELECT distinguishedName, description, comment, info FROM 'LDAP://gmzrzdc001.ferchau.local:3268' WHERE objectCategory='person' and objectClass = 'user'" 
		Set objRecordSet = objCommand.Execute
	objRecordSet.MoveFirst

	Do Until objRecordSet.EOF
		strdistinguishedName =  objRecordSet.Fields("distinguishedName").Value 
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
			'For Each strLine In arrComment
				strComment = "'" & Replace(strLine,"'","`") & "'"
			'Next
		End If
		arrInfo = objRecordSet.Fields("info").Value
		If IsNull(arrInfo) Then
			strInfo = "NULL"
		Else
			'For Each strLine In arrInfo
				strInfo = "'" & Replace(strLine,"'","`") & "'"
			'Next
		End If
		IF db.State = 1 then
			strSQL = "update [IAM].[dbo].[t_AD_User] set description = " & strDescr & ", comment = " & strComment & ", info = " & strInfo & " where distinguishedName = '" & strdistinguishedName & "'"
			db.Execute(strSQL)
			IF Err.Number <> 0 Then
				db.Execute("insert into dbo.t_logfile(Meldungstext) values('Error # " & CStr(Err.Number) & " - " & Err.Description & "')")
				Err.Clear
			End If
		End If
		objRecordSet.MoveNext
	Loop
Else
	IF Err.Number <> 0 Then
		db.Execute("insert into dbo.t_logfile(Meldungstext) values('Error # " & CStr(Err.Number) & " - " & Err.Description & "')")
		Err.Clear
	End If
End If	
db.close
objConnection.close
