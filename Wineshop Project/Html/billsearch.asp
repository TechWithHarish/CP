<%
Dim rs, billno, db
billno = Request.QueryString("billno")

If billno <> "" Then
    ' Create connection and recordset objects

    Set db = Server.CreateObject("ADODB.Connection")
    Set rs = Server.CreateObject("ADODB.Recordset")
    
    ' Set up database connection

    On Error Resume Next
    db.Provider = "Microsoft.Jet.OLEDB.4.0"
    db.Open "H:\Wineshop Project\Wineshop_Database_1.mdb"
    
    If Err.Number <> 0 Then
        Response.Write "<p>Error connecting to database: " & Err.Description & "</p>"
        Err.Clear
        Set rs = Nothing
        Set db = Nothing
        Response.End
    End If
    On Error GoTo 0

    ' Open the recordset
    rs.Open "orders_details", db, 1, 3

    ' Check for the existence of the 'OrderID' field

    Dim fieldExists
    fieldExists = False
    For i = 0 To rs.Fields.Count - 1
        If LCase(rs.Fields(i).Name) = "orderid" Then
            fieldExists = True
            Exit For
        End If
    Next

    If fieldExists Then
        ' Filter the recordset for the provided bill number

        rs.Filter = "OrderID = '" & billno & "'"
        
        If rs.EOF Then
            Response.Write "<p>No bill found with Bill No: " & Server.HTMLEncode(billno) & "</p>"
        Else
            ' Print the Bill No, Customer Name, and Customer Address outside the table

	Response.Write "<div style='text-align: left;'>"
	Response.Write "<p style='color: white;'><strong>Bill No:</strong> " & Server.HTMLEncode(rs("OrderID")) & "</p>"
	Response.Write "<p style='color: white;'><strong>Customer Name:</strong> " & Server.HTMLEncode(rs("CustomerName")) & "</p>"
	Response.Write "<p style='color: white;'><strong>Customer Address:</strong> " & Server.HTMLEncode(rs("CustomerAddress")) & "</p>"
	Response.Write "</div>"

            ' Generate the HTML table for displaying order details

            Response.Write "<table border='1' style='width: 100%; border-collapse: collapse;'>"
            Response.Write "<thead><tr>"
            Response.Write "<th>Product</th>"
            Response.Write "<th>Quantity</th>"
            Response.Write "<th>Price</th>"
            Response.Write "<th>Total Amount</th>"
            Response.Write "</tr></thead>"
            Response.Write "<tbody>"

            ' Assuming the products, quantities, and prices are stored as concatenated strings
            Dim productsArray, quantitiesArray, pricesArray
            productsArray = Split(rs("Products"), ", ")
            quantitiesArray = Split(rs("Quantities"), ", ")
            pricesArray = Split(rs("Prices"), ", ")

            Dim j, totalAmount, productTotalAmount
            totalAmount = 0 ' Initialize the total amount for the entire order

            For j = 0 To UBound(productsArray)
                Response.Write "<tr>"
                
                ' Calculate the total amount for each product

                productTotalAmount = CDbl(pricesArray(j)) * CInt(quantitiesArray(j))
                totalAmount = totalAmount + productTotalAmount ' Accumulate total amount

                Response.Write "<td>" & Server.HTMLEncode(productsArray(j)) & "</td>"
                Response.Write "<td>" & Server.HTMLEncode(quantitiesArray(j)) & "</td>"
                Response.Write "<td>$" & FormatNumber(CDbl(pricesArray(j)), 2) & "</td>"
                Response.Write "<td>$" & FormatNumber(productTotalAmount, 2) & "</td>" ' Display product total
                Response.Write "</tr>"
            Next

            Response.Write "</tbody></table>"
            Response.Write "<h3>Total Amount for All Products: $" & FormatNumber(totalAmount, 2) & "</h3>" ' Display the total amount for the order

            ' Add a button to go back to the homepage

            Response.Write "<div style='margin-top: 20px;'>"
            Response.Write "<a href='../Html/wineshop_index.html' style='text-decoration: none;'>" ' Link to your homepage
            Response.Write "<button style='padding: 10px 20px; background-color: #4CAF50; color: white; border: none; border-radius: 5px;'>Go to Homepage</button>"
            Response.Write "</a>"
            Response.Write "</div>"

            ' Link to the CSS file for styling
            Response.Write "<link rel='stylesheet' type='text/css' href='../Css/invoice_confirmation.css'>" ' Adjust the path to your CSS file

        End If
    Else
        Response.Write "<p>Error: 'OrderID' field not found in the table.</p>"
    End If

    ' Clean up

    rs.Close
    db.Close
    Set rs = Nothing
    Set db = Nothing
Else
    Response.Write "<p>Please provide a bill number.</p>"
End If
%>
