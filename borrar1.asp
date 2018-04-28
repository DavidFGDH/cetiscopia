<HTML>
<HEAD>
<TITLE>Borrar1.asp</TITLE>
</HEAD>
<BODY>
<div align="center">
<h1>Borrar un registro</h1>
<br>
<%		
'Instanciamos y abrimos nuestro objeto conexion 
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open "Mibase"
%>

<FORM METHOD="POST" ACTION="borrar2.asp">
Nombre<br>
<%
'Creamos la sentencia SQL y la ejecutamos
sSQL="Select nombre From clientes Order By nombre"
set RS = conn.execute(sSQL)
%>
<select name="nombre">
<%
'Generamos el menu desplegable
do while not RS.eof%>
	<option><%=RS("nombre")%>
	<%RS.movenext
	loop
%>
</select>
<br>
<INPUT TYPE="SUBMIT" value="Borrar">
</FORM>
</div>

</BODY>
</HTML>
