<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Invoice</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            background-color: #f4f4f4;
            margin: 0;
            padding: 20px;
        }

        h2 {
            text-align: center;
            color: #333;
        }

        p {
            font-size: 1.1em;
            color: #555;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            background-color: #fff;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
        }

        th, td {
            padding: 10px;
            text-align: left;
            border: 1px solid #ddd;
        }

        th {
            background-color: #4CAF50;
            color: white;
        }

        tr:nth-child(even) {
            background-color: #f9f9f9;
        }

        tr:hover {
            background-color: #f1f1f1;
        }

        strong {
            color: #333;
        }

        .button {
            background-color: #4CAF50;
            color: white;
            padding: 10px 15px;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            font-size: 1em;
            display: block;
            width: 150px;
            margin: 20px auto;
            text-align: center;
            text-decoration: none;
        }

        .button:hover {
            background-color: #45a049;
        }

        .thank-you {
            text-align: center;
            font-size: 1.2em;
            color: #333;
            margin-top: 30px;
        }

        @media print {
            body * {
                visibility: hidden;
            }

            table, table * {
                visibility: visible;
            }

            table {
                position: absolute;
                left: 0;
                top: 0;
            }
        }
    </style>
</head>
<body>

    <%
      ' Error handling block
      On Error Resume Next

      ' Get the Customer ID from the query string
      customerID = request.querystring("cid")

      ' Create ADO connection and recordset objects
      set db = server.createobject("adodb.connection")
      set rs = server.createobject("adodb.recordset")

      ' Open the database connection
      db.provider = "microsoft.jet.oledb.4.0"
      db.open "H:\Wineshop Project\occ\DATABASE\Beauty2 - Copy.mdb"

      ' Query the table to get the order details for the given Customer ID
      sqlQuery = "SELECT * FROM tableBeauty WHERE CID = " & customerID
      rs.open sqlQuery, db, 1, 3 ' Open recordset to retrieve the order details

      ' Initialize total variable
      grandTotal = 0

      ' Check if the record exists
      if not rs.eof then
        ' Get the customer and order details
        customerName = rs(1)
        email = rs(2)
        services = Split(rs(3), ",") ' Assuming services are stored as a comma-separated string
        total = Split(rs(4), ",") ' Assuming rates are stored as a comma-separated string
        preferredDate = rs(5)
        preferredTime = rs(6)

        ' Calculate the grand total
        For i = LBound(total) To UBound(total)
          grandTotal = grandTotal + CDBL(total(i)) ' Convert rate to double for accurate total
        Next
      end if

      ' Close the recordset and database connection
      rs.close
      db.close
      set rs = nothing
      set db = nothing

      ' Handle any errors
      If Err.Number <> 0 Then
          response.write("<br><br>Error occurred: " & Err.Description)
      End If
    %>

    <!-- Display the Bill -->
    <h2>Bill for Customer: <%= customerName %></h2>
    <p>Email: <%= email %></p>
    <p>Selected Services: <%= Join(services, ", ") %></p>
    <p>Total Rates: <%= Join(total, ", ") %></p>
    <p>Preferred Date: <%= preferredDate %></p>
    <p>Preferred Time: <%= preferredTime %></p>

    <h3>Summary</h3>
    <table>
      <tr>
        <th>Description</th>
        <th>Details</th>
      </tr>
      <tr>
        <td>Customer Name</td>
        <td><%= customerName %></td>
      </tr>
      <tr>
        <td>Email</td>
        <td><%= email %></td>
      </tr>
      <tr>
        <td>Selected Services</td>
        <td><%= Join(services, ", ") %></td>
      </tr>
      <tr>
        <td>Total Rates</td>
        <td><%= Join(total, ", ") %></td>
      </tr>
      <tr>
        <td>Preferred Date</td>
        <td><%= preferredDate %></td>
      </tr>
      <tr>
        <td>Preferred Time</td>
        <td><%= preferredTime %></td>
      </tr>
      <tr>
        <td><strong>Grand Total</strong></td>
        <td><strong><%= FormatCurrency(grandTotal) %></strong></td>
      </tr>
    </table>

    <h3 class="thank-you">Thank you for placing your order!</h3>
    <button class="button" onclick="window.print()">Print Bill</button>
    
</body>
</html>
