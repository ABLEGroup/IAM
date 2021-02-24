Dim  Input
Dim strRecset, strLine, strDescr, strSQL
Dim arrDescr

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
	db.Execute("truncate table [IAM].[dbo].[t_AD_Gruppen]")
	IF Err.Number <> 0 Then
		db.Execute("insert into dbo.t_logfile(Meldungstext) values('Error # " & CStr(Err.Number) & " - " & Err.Description & "')")
		Err.Clear
	End If

	objCommand.CommandText = _
		"SELECT samaccountname,mail,sn,name,cn,objectCategory,distinguishedName,whenChanged,whenCreated,description,objectSID FROM 'LDAP://gmzrzdc001.ferchau.local:3268' WHERE objectCategory='group'" 
		Set objRecordSet = objCommand.Execute
	objRecordSet.MoveFirst

	Do Until objRecordSet.EOF
		strRecset = "[IAM].[dbo].[f_get_Domain]('" & objRecordSet.Fields("distinguishedName").Value & "'),'" & objRecordSet.Fields("name").Value & "','" & objRecordSet.Fields("samaccountname").Value & "','" & objRecordSet.Fields("cn").Value & "','" & objRecordSet.Fields("mail").Value _
			& "','" & objRecordSet.Fields("distinguishedName").Value  & "','" & objRecordSet.Fields("objectCategory").Value  & "',cast('" & objRecordSet.Fields("whenChanged").Value  & "' as smalldatetime),cast('" & objRecordSet.Fields("whenCreated").Value & "' as smalldatetime)"

		arrDescr = objRecordSet.Fields("description").Value
		If IsNull(arrDescr) Then
			strDescr = "''"
		Else
			For Each strLine In arrDescr
				strDescr = "'" & Replace(strLine,"'","`") & "'"
			Next
		End If
		strRecset = strRecset  & "," & strDescr 
		IF db.State = 1 then
			strSQL = "insert INTO [IAM].[dbo].[t_AD_Gruppen]([DC],[GruppenName],[samaccountname],[cn],[mail],[distinguishedName],[objectCategory],[whenChanged],[whenCreated],[description]) values (" & strRecset & ")"
			'files.WriteLine(strSQL)
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
