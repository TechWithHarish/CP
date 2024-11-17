<html>
  <body>
    <%
      ' Error handling block
      On Error Resume Next

      ' Create ADO connection and recordset objects
      set db = server.createobject("adodb.connection")
      set rs = server.createobject("adodb.recordset")

      ' Open the database connection
      db.provider = "microsoft.JET.oledb.4.0"
      db.open "H:\Wineshop Project\occ\DATABASE\Beauty2 - Copy.mdb"

      ' Open the recordset (adOpenKeyset for moving through records, adLockOptimistic for adding new records)
      rs.open "tableBeauty", db, 1, 3

      ' Display success message for order placement
      response.write("<br><br>")
      response.write("<b>Your Order has been placed successfully!</b>")
      response.write("<br><br>")

      ' Initialize customer ID (CID)
      if rs.eof and rs.bof then
        CID = 1001 ' Start at 1001 if there are no records
      else
        rs.movelast ' Move to the last record to get the last customer ID
        CID = rs(0) + 1 ' Increment the last ID by 1
      end if

      ' Get the grand total from the form
      Dim grandTotal
      grandTotal = CDbl(request.form("rates")) ' Get the total rates submitted from the form

      ' Check if the total is valid
      If grandTotal <= 0 Then
          response.write("<br><br>Error: Invalid total rates.")
          response.end ' Stop further processing
      End If

      ' Add a new record to the table
      rs.addnew()
        rs(0) = CID ' Set Customer ID
        rs(1) = request.form("n") ' Set Customer Name
        rs(2) = request.form("e") ' Set Email
        rs(3) = request.form("ss") ' Set Selected Services (can be modified to store multiple values if needed)
        rs(4) = grandTotal ' Set Grand Total
        rs(5) = request.form("d") ' Set Preferred Date
        rs(6) = request.form("t") ' Set Preferred Time
      rs.update()

      ' Redirect to bill page with Customer ID as a query string
      response.redirect "generateBill.asp?cid=" & CID

      ' Close the recordset and database connection
      rs.close
      db.close
      set rs = nothing

      ' Handle any errors
      If Err.Number <> 0 Then
          response.write("<br><br>Error occurred: " & Err.Description)
      End If

      ' End error handling block
      On Error GoTo 0
    %>
  </body>
</html>
