 <% 
dim con 
dim rs 

Set Con = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.RecordSet")
Set Rs1 = Server.CreateObject("ADODB.RecordSet")
Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	Con.ConnectionString = "PROVIDER=SQLOLEDB;DATA SOURCE=192.168.0.148;UID=macromis;PWD=Mis@0000;DATABASE=MIS"
%>

	