var connection = new ActiveXObject("ADODB.Connection") ;
var password = "abcd";
var connectionstring="Data Source=aaa.aaa.com;Initial Catalog=<catalog>;User ID=test;password;Provider=SQLOLEDB";

connection.Open(connectionstring);
var rs = new ActiveXObject("ADODB.Recordset");

rs.Open("SELECT * FROM table", connection);
rs.MoveFirst
while(!rs.eof)
{
   document.write(rs.fields(1));
   rs.movenext;
}

rs.close;
connection.close; 
