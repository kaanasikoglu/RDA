<html>
<head>
<title>Update Entry Select</title>
<style>
body{
			background-color: #282828;
			color: #e7e7e7;
			font-family: Arial;
			padding: 20px;
		}
		a{
			color: green;
			font-family: Arial;
		}
</style>
</head>
<body>
<%
'Dimension variables
Dim adoCon 			'Holds the Database Connection Object
Dim rsFeedback		'Holds the recordset for the records in the database
Dim strSQL			'Holds the SQL query for the database

'Create an ADO connection odject
Set adoCon = Server.CreateObject("ADODB.Connection")

'Set an active connection to the Connection object using a DSN-less connection
adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("feedback.mdb")

'Create an ADO recordset object
Set rsFeedback = Server.CreateObject("ADODB.Recordset")

'Initialise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT tblComments.* FROM tblComments;"

'Open the recordset with the SQL query 
rsFeedback.Open strSQL, adoCon

'Loop through the recordset
Do While not rsFeedback.EOF
	
	'Write the HTML to display the current record in the recordset
	Response.Write ("<br>")
	Response.Write ("<a href=""update_form.asp?ID=" & rsFeedback("ID") & """>")
	Response.Write (rsFeedback("Name")) 
	Response.Write ("</a>")
	Response.Write ("<br>")
	Response.Write (rsFeedback("Email"))
	Response.Write ("<br>")

	Response.Write (rsFeedback("Message"))
	Response.Write ("<br>")

	'Move to the next record in the recordset
	rsFeedback.MoveNext

Loop

'Reset server objects
rsFeedback.Close
Set rsFeedback = Nothing
Set adoCon = Nothing
%>
</body>
</html>