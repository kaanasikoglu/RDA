<html>
	<head>
		<title>Feedback</title>
	<style>
	body{
		background-color: #282828;
		color: #e7e7e7;
		font-family: Arial;
		padding: 20px;
	}
	</style>
	</head>
	<body>
	<%
		Dim adoCon 			'Holds the Database Connection Object
		Dim rsFeedback		'Holds the recordset for the records in the database
		Dim strSQL			'Holds the SQL query for the database
		Set adoCon = Server.CreateObject("ADODB.Connection")
		adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("feedback.mdb")
		Set rsFeedback = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT tblComments.ID, tblComments.Name, tblComments.Email, tblComments.Message FROM tblComments;"
		rsFeedback.Open strSQL, adoCon
		Do While not rsFeedback.EOF
			Response.Write ("<br>ID: ")
			Response.Write (rsFeedback("ID"))
			Response.Write ("<br>Name: ")
			Response.Write (rsFeedback("Name"))
			Response.Write ("<br>E-mail: ")
			Response.Write (rsFeedback("Email"))
			Response.Write ("<br> Message: ")
			Response.Write (rsFeedback("Message"))
			Response.Write ("<br>")
			rsFeedback.MoveNext
			Loop
		rsFeedback.Close
		Set rsFeedback = Nothing
		Set adoCon = Nothing
	%>
	</body>
</html>