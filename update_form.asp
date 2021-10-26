<%
'Dimension variables
Dim adoCon 			'Holds the Database Connection Object
Dim rsFeedback		'Holds the recordset for the record to be updated
Dim strSQL			'Holds the SQL query for the database
Dim lngRecordNo		'Holds the record number to be updated

'Read in the record number to be updated
lngRecordNo = CLng(Request.QueryString("ID"))

'Create an ADO connection odject
Set adoCon = Server.CreateObject("ADODB.Connection")

'Set an active connection to the Connection object using a DSN-less connection
adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("feedback.mdb")

'Create an ADO recordset object
Set rsFeedback = Server.CreateObject("ADODB.Recordset")

'Initialise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT tblComments.* FROM tblComments WHERE ID=" & lngRecordNo

'Open the recordset with the SQL query 
rsFeedback.Open strSQL, adoCon
%>
<html>
<head>
<title>Feedback Update Form</title>
<style>
	body{
			background-color: #282828;
			color: #e7e7e7;
			font-family: Arial;
			padding: 20px;
		}
</style>
</head>
<body bgcolor="white" text="black">
<!-- Begin form code -->
<form name="form" method="post" action="update_entry.asp">
  Name: <input type="text" name="name" maxlength="20" value="<% = rsFeedback("Name") %>">
  <br><br>
  Email: <input type="text" name="email" maxlength="20" value="<% = rsFeedback("Email") %>">
  <br><br>
  Comments: <input type="text" name="message" maxlength="60" value="<% = rsFeedback("Message") %>">
  <input type="hidden" name="ID" value="<% = rsFeedback("ID") %>">
  <br><br>
  <input type="submit" name="Submit" value="Submit">
</form>
<!-- End form code -->
</body>
</html>
<%
'Reset server objects
rsFeedback.Close
Set rsFeedback = Nothing
Set adoCon = Nothing
%>
