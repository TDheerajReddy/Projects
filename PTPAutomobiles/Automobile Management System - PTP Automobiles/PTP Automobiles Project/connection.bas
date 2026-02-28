Attribute VB_Name = "Connection"
Public dbconn As New ADODB.Connection
Public recset As New ADODB.Recordset

Public Sub dbconnection1()
    dbconn.Open "provider=MSDASQL;driver={Mysql ODBC 3.51 Driver};database=PTPAutomobiles;server=localhost;user=root;password=123"
    recset.Open "select * from OrderToManufacturer", dbconn, adOpenDynamic, adLockOptimistic
End Sub
Public Sub dbconnection2()
    dbconn.Open "provider=MSDASQL;driver={Mysql ODBC 3.51 Driver};database=PTPAutomobiles;server=localhost;user=root;password=123"
    recset.Open "select * from OrderFromCustomer", dbconn, adOpenDynamic, adLockOptimistic
End Sub
Public Sub dbconnection3()
    dbconn.Open "provider=MSDASQL;driver={Mysql ODBC 3.51 Driver};database=PTPAutomobiles;server=localhost;user=root;password=123"
    recset.Open "select * from Servicing", dbconn, adOpenDynamic, adLockOptimistic
End Sub
Public Sub dbconnection4()
    dbconn.Open "provider=MSDASQL;driver={Mysql ODBC 3.51 Driver};database=PTPAutomobiles;server=localhost;user=root;password=123"
    recset.Open "select * from SpareParts", dbconn, adOpenDynamic, adLockOptimistic
End Sub
Public Sub dbconnection5()
    dbconn.Open "provider=MSDASQL;driver={Mysql ODBC 3.51 Driver};database=PTPAutomobiles;server=localhost;user=root;password=123"
    recset.Open "select * from Employee", dbconn, adOpenDynamic, adLockOptimistic
End Sub
Public Sub closeDB()
    Set dbconn = Nothing
    Set recset = Nothing
End Sub
