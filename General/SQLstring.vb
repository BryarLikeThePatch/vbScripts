'ArrayToString

Function arrayToString(arr As Variant()) As String
    Dim s As String
    For Each v in arr
        s = s & "," & v
    Next v
    arrayToString = Mid(s,2,Len(s)) 'returns string minus leading comma
End Function 

'this code is especially helpful if you need to create an SQL "WHERE ... IN... " filter string. For example:
'[  ][    A    ][    B     ][    C    ]
'[ 1][ cust_No ][ cust_Nme ][ sales_$ ]
'[ 2][      123][     name1][      452]
'[ 3][      124][     name2][    1,000]
'[ 4][      125][     name3][   10,000]
'[ 5][      126][     name4][      631]
'[ 6][      127][     name5][    1,001]
'
'
'using the FILTER() function would return an array based on sales dollars over 1000
'=FILTER(A:A,C:C>=1000) //[124],[125],[127] 

'this would create a spill array with each value from the FILTER() function getting its own cell

'Alternatively, wrapping the above FILTER() in the arrayToString() function would return a string to a single cell. 
'=arrayToString(FILTER(A:A,C:C>=1000)) //"124,125,127"

'Another variation of this same function to be used for SQL couuld be as follows:

Function SQLSTRING(arr As Variant()) As String
    Dim s as String
    For Each v in arr
        s = s & "','" & v
    Next v
    SQLSTRING = Mid(s,3,Len(s)) & "'"
End Function 

'using the above FILTER() example. the returned value is as follows:
'=SQLSTRING(FILTER(A:A,C:C>=1000)) //"'124','125','127'" 

'This is helpful for a situation where the SQL Statement is looking for specifically those customers:
'
'SELECT
'   customerNumber
'   ,customerName
'   ,customerCreatecDate
'   ,ytdSales
'   ,clientContact
'   ,clientContactEmail
'FROM
'   customer
'       INNER JOIN contact ON customer.id = contact.regardingCustomerID
'WHERE
'   customerNumber IN ('124','125','127');