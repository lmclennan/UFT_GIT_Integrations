
SystemUtil.Run "C:\Users\chrismclennan\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Firefox.lnk","","",""
Browser("Browser").Navigate "advantageonlineshopping.com" @@ hightlight id_;_919610_;_script infofile_;_ZIP::ssf1.xml_;_
Browser("Browser").Close

Sub test1

	strExpFirstName = "testy2"
	
	'1 initiate the database connection
	Set oDB = CreateObject("ADODB.Connection")
	
	'2 define the connection string and open the database
	oDB.ConnectionString = "Provider=SQLOLEDB;Server=CHRISMCLENNB79A;Database=master;user id=uft;password=fut"
	oDB.Open
	
	'3 define the sql query
	'SQL="select * from student_list where StudentID='2'"
	SQL="select * from student_list where Year='Freshman' and StudentID='3'"
	
	'4 create a blank recordset object
	Set oRec = CreateObject("ADODB.Recordset")
	
	'5 open the recordset by running the sql query in the opened database
	oRec.Open SQL,oDB
	
	'6 validate the data displayed in recordset
	If oRec.EOF = False Then
		'msgbox "record returned"
		strActFirstName = oRec.Fields("FirstName").Value
		strActLastName = oRec.Fields("LastName").Value
		
		If strExpFirstName = strActFirstName Then
			msgbox "first name matched"
		Else
			msgbox "first name not matched" & vbcrlf & "expected: " & strExpFirstName & "  actual: " & strActFirstName
		End If
		
	Else
		msgbox "no record returned"
	End If
	
	'7 close the recordset object and databse
	oRec.Close
	oDB.Close
	
	'DbTable("DbTable").Check CheckPoint("DbTable")
	'DbTable("DbTable_2").Check CheckPoint("DbTable_2")
End Sub

Sub test2
	
	Set objconn=CreateObject("ADODB.connection")
	Set objrecordset=CreateObject("ADODB.Recordset")
	
	objconn.Provider=("Microsoft.ACE.OLEDB.12.0")
	objconn.Open "C:\Users\chrismclennan\Documents\Database1.accdb"
	objrecordset.Open "Select FirstName,LastName from Students",objconn
	
	Do Until objrecordset.EOF=true
	msgbox objrecordset.Fields("FirstName")
	msgbox objrecordset.Fields("LastName")
	objrecordset.MoveNext
	Loop
	
	objrecordset.Close
	objconn.Close
	
	Set objrecordset=Nothing
	Set objconn=Nothing
End Sub
