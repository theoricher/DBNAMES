strComputer = "localhost" 

Set objConnection = CreateObject("ADODB.Connection")

objConnection.Open _
    "Provider=SQLOLEDB;Data Source=" & strComputer & ";" & _
        "Trusted_Connection=Yes;Initial Catalog=Master"

Set objRecordset = objConnection.Execute("Select Name, filename From SysDatabases")
If objRecordset.Recordcount = 0 Then
    Wscript.Echo ""
Else
	
    Do Until objRecordset.EOF
		Wscript.Echo "<DBNAMES>"
        Wscript.Echo "<NAME>" & objRecordset.Fields("Name") & "</NAME>"
		Wscript.Echo "<FILE>" & objRecordset.Fields("filename") &  "</FILE>"
        objRecordset.MoveNext
		rem objFilename.MoveNext
		Wscript.Echo "</DBNAMES>"
    Loop
	

End If

