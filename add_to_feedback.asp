<%
Dim adoCon 			'Holds the Database Connection Object
Dim rsAddComments	'Holds the recordset for the new record to be added to the database
Dim strSQL			'Holds the SQL query for the database


Set adoCon = Server.CreateObject("ADODB.Connection")


adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("feedback.mdb")


Set rsAddComments = Server.CreateObject("ADODB.Recordset")


strSQL = "SELECT tblComments.Name, tblComments.Email, tblComments.Message FROM tblComments;"


rsAddComments.CursorType = 2


rsAddComments.LockType = 3


rsAddComments.Open strSQL, adoCon


rsAddComments.AddNew


rsAddComments.Fields("Name") = Request.Form("name")
rsAddComments.Fields("Email") = Request.Form("email")
rsAddComments.Fields("Message") = Request.Form("message")


rsAddComments.Update


rsAddComments.Close
Set rsAddComments = Nothing
Set adoCon = Nothing

Response.Redirect "feedback.asp"
%>