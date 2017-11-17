var connection = new ActiveXObject("ADODB.Connection") ;

/*var connectionstring="Data Source=ITS-LIMS2;Initial Catalog=<catalog>;User ID=<user>;Password=<password>;Provider=SQLOLEDB";*/
var connectionstring="Data Source=ITS-LIMS\PROD;Initial Catalog=ITS_WFLIMS;Integrated Security=True";
connection.Open(connectionstring);
var rs = new ActiveXObject("ADODB.Recordset");

rs.Open("SELECT * FROM WORK_ORDERS", connection);
rs.MoveFirst
while(!rs.eof)
{
   document.write(rs.fields(1));
   rs.movenext;
}

rs.close;
connection.close;