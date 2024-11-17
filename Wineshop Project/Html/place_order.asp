<%
Set db = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
db.Provider = "Microsoft.Jet.OLEDB.4.0"
db.Open "H:\Wineshop Project\Wineshop_Database_1.mdb"
rs.Open "orders_details", db, 1, 3

if rs.eof and rs.bof then
    x = 1001
else
    rs.MoveLast
    

    If IsNull(rs(0)) Then
        x = 1001 ' Start at 1001 if the value is null
    Else
        ' Extract the last 3 characters and convert to an integer
        z = Right(CStr(rs(0)), 4) 

        ' Ensure the extracted part is numeric before converting
        If IsNumeric(z) Then
            x = CInt(z) + 1 ' Increment the numeric part by 1
        Else
            x = 1001 
        End If
    End If
end if

Dim customerName, customerAddress
customerName = Server.HTMLEncode(Request.Form("customer_name"))
customerAddress = Server.HTMLEncode(Request.Form("customer_address"))

Dim productsList, quantitiesList, pricesList, totalAmount
totalAmount = 0
productsList = ""
quantitiesList = ""
pricesList = ""

Dim i, currentProductName, currentProductPrice, currentProductQuantity, currentTotal
For i = 1 To 8
    currentProductName = Server.HTMLEncode(Request.Form("product_name_" & i))
    If currentProductName <> "" Then
        currentProductPrice = CDbl(Request.Form("product_price_" & i))

        Dim quantityInput
        quantityInput = Request.Form("quantity_" & i)

        ' Validate and convert the quantity input
        If IsNumeric(quantityInput) Then
            currentProductQuantity = CLng(quantityInput) ' Use CLng for larger numbers
        Else
            currentProductQuantity = 0 ' Default to 0 if not a number
        End If

        currentTotal = currentProductPrice * currentProductQuantity

        If currentProductQuantity > 0 Then
            If productsList <> "" Then
                productsList = productsList & ", "
                quantitiesList = quantitiesList & ", "
                pricesList = pricesList & ", "
            End If

            productsList = productsList & currentProductName
            quantitiesList = quantitiesList & currentProductQuantity
            pricesList = pricesList & currentProductPrice

            totalAmount = totalAmount + currentTotal
        End If
    End If
Next

rs.AddNew
rs("OrderID") = "OD" & x ' Assign the generated primary key value here
rs("CustomerName") = customerName
rs("CustomerAddress") = customerAddress
rs("Products") = productsList
rs("Quantities") = quantitiesList
rs("Prices") = pricesList
rs("TotalAmount") = totalAmount
rs.Update ' Ensure all required fields are set before calling Update

' Generate the Invoice HTML
Dim invoiceHtml
invoiceHtml = "<html>" & _
              "<head>" & _
              "<title>Invoice</title>" & _
              "<link rel='stylesheet' href='../Css/invoice_confirmation.css'>" & _
              "</head>" & _
              "<body>" & _
              "<h1>Invoice</h1>" & _
	      "<h3><strong>Bill No:</strong> " & (rs("OrderID")) & "</h3>" & _
              "<h3><strong>Customer Name:</strong> " & customerName & "</h3>" & _
              "<h3><strong>Customer Address:</strong> " & customerAddress & "</h3>" & _
              "<table border='1' style='width: 100%; border-collapse: collapse;'>" & _
              "<tr><th>Product Name</th><th>Quantity</th><th>Price</th></tr>"

Dim productArray, quantityArray, priceArray
productArray = Split(productsList, ", ")
quantityArray = Split(quantitiesList, ", ")
priceArray = Split(pricesList, ", ")

Dim j
For j = 0 To UBound(productArray)
    invoiceHtml = invoiceHtml & "<tr>" & _
                  "<td>" & productArray(j) & "</td>" & _
                  "<td>" & quantityArray(j) & "</td>" & _
                  "<td>$" & priceArray(j) & "</td>" & _
                  "</tr>"
Next

invoiceHtml = invoiceHtml & "</table>" & _
              "<h3><strong>Total Amount:</strong> $" & totalAmount & "</h3>" & _
              "</body></html>"

Response.Write(invoiceHtml)

' Generate the Confirmation HTML
Dim confirmationHtml
confirmationHtml = "<link rel='stylesheet' href='../Css/invoice_confirmation.css'>" & _
                   "<div class='confirmation-container'>" & _
                   "<h2>Thank You! Your Order Has Been Placed.</h2>" & _
                   "<p>Your order details have been successfully recorded in our system.</p>" & _
                   "<a href='../Html/wineshop_index.html' class='back-link'>Go to Homepage</a>" & _
                   "</div>"

Response.Write(confirmationHtml)

rs.Close
db.Close
Set rs = Nothing
Set db = Nothing
%>
