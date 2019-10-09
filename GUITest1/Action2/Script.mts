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

