<%
set db = server.createobject("adodb.connection")
set rs = server.createobject("adodb.recordset")

db.provider = "microsoft.jet.oledb.4.0"
db.open "H:\Wineshop Project\Wineshop_Database_1.mdb"
rs.open "contact_details", db, 1, 3

if rs.eof and rs.bof then
    ' No records exist, start the sequence at 101
    x = 101
else
    ' Move to the last record to get the most recent entry
    rs.MoveLast
    
    ' Check if rs(0) is Null, handle safely
    If IsNull(rs(0)) Then
        x = 101 ' Start at 101 if the value is null
    Else
        ' Extract the last 3 characters and convert to an integer
        z = Right(CStr(rs(0)), 3) 

        ' Ensure the extracted part is numeric before converting
        If IsNumeric(z) Then
            x = CInt(z) + 1 ' Increment the numeric part by 1
        Else
            x = 101 ' Reset to 101 if the value is not numeric
        End If
    End If
end if

' Add new record
rs.AddNew()
    rs(0) = "CON" & x ' Create the new contact ID
    rs(1) = request.form("name") ' Capture form data
    rs(2) = request.form("email")
    rs(3) = request.form("message")
rs.Update()

' Close the recordset and the connection
rs.close()
db.close()

' Create the confirmation HTML
dim confirmationHtml
confirmationHtml = "<link rel='stylesheet' href='../Css/invoice_confirmation.css'>"
confirmationHtml = confirmationHtml & "<div class='confirmation-container' style=""margin-top:170px;"">"
confirmationHtml = confirmationHtml & "<h2>Thank You! Your Message Has Been Received.</h2>"
confirmationHtml = confirmationHtml & "<p>Your contact details have been successfully recorded in our system. We will get back to you shortly.</p>"
confirmationHtml = confirmationHtml & "<a href='../Html/wineshop_index.html' class='back-link'>Go to Homepage</a>"
confirmationHtml = confirmationHtml & "</div>"

' Output the confirmation message
Response.Write(confirmationHtml)
%>
</body>
</html>
