<%
' Declare variables
Dim variableName

' Set up the page template and header
Response.Write "<!DOCTYPE HTML PUBLIC ""-//IETF//DTD HTML 2.0//EN"">" & vbCrLf
Response.Write "<html>" & vbCrLf
Response.Write vbTab & "<head>" & vbCrLf
Response.Write vbTab & vbTab & "<meta http-equiv=""Content-Type"" content=""text/html; charset=ISO-8859-1"">" & vbCrLf
Response.Write vbTab & vbTab & "<title>ASP Request Collections</title>" & vbCrLf
Response.Write vbTab & "</head>" & vbCrLf
Response.Write vbTab & "<body>" & vbCrLf
Response.Write vbTab & vbTab & "<h1>ASP Request Collections</h1>" & vbCrLf

' Display server variables
Response.Write vbTab & vbTab & "<h2>Server Variables</h2>" & vbCrLf
For Each variableName In Request.ServerVariables
	Response.Write vbTab & vbTab & "<strong>" & variableName & "</strong><br><xmp>" & Request.ServerVariables(variableName) & "</xmp><br>" & vbCrLf
Next
Response.Write vbTab & vbTab & "<hr>" & vbCrLf

' Display query string variables
Response.Write vbTab & vbTab & "<h2>Query String</h2>" & vbCrLf
For Each variableName In Request.QueryString
	Response.Write vbTab & vbTab & "<strong>" & variableName & "</strong><br><xmp>" & Request.QueryString(variableName) & "</xmp><br>" & vbCrLf
Next
Response.Write vbTab & vbTab & "<hr>" & vbCrLf

' Display form variables
Response.Write vbTab & vbTab & "<h2>Form</h2>" & vbCrLf
For Each variableName In Request.Form
	Response.Write vbTab & vbTab & "<strong>" & variableName & "</strong><br><xmp>" & Request.Form(variableName) & "</xmp><br>" & vbCrLf
Next
Response.Write vbTab & vbTab & "<hr>" & vbCrLf

' Display session variables
Response.Write vbTab & vbTab & "<h2>Session</h2>" & vbCrLf
For Each variableName In Session.Contents
	Response.Write vbTab & vbTab & "<strong>" & variableName & "</strong><br><xmp>" & Session(variableName) & "</xmp><br>" & vbCrLf
Next
Response.Write vbTab & vbTab & "<hr>" & vbCrLf

' Display cookies
Response.Write vbTab & vbTab & "<h2>Cookies</h2>" & vbCrLf
For Each variableName In Request.Cookies
	Response.Write vbTab & vbTab & "<strong>" & variableName & "</strong><br><xmp>" & Request.Cookies(variableName) & "</xmp><br>" & vbCrLf
Next
Response.Write vbTab & vbTab & "<hr>" & vbCrLf

' Display the page footer
Response.Write vbTab & vbTab & "<p>Valid HTML 2.0.</p>" & vbCrLf
Response.Write vbTab & "</body>" & vbCrLf
Response.Write "</html>" & vbCrLf
%>