<HTML>
<HEAD>
<TITLE>Lectura de registros de una tabla</TITLE>
</HEAD>
<BODY>
<h1><div align="center">Lectura de la tabla</div></h1>
<br>
<br>
<%
'Antes de nada hay que instanciar el objeto Connection 
Set Conn = Server.CreateObject("ADODB.Connection")

'Una vez instanciado Connection lo podemos abrir y le asignamos la base de datos donde vamos a efectuar las operaciones
Conn.Open "Mibase"

'Ahora creamos la sentencia SQL que nos servira para hablar a la BD
sSQL="Select * From Clientes Order By nombre"

'Ejecutamos la orden 
set RS = Conn.Execute(sSQL)

'Mostramos los registros%>
<table align="center">
<tr>
<th>Nombre</th>
<th>Teléfono</th>
</tr>
<%
Do While Not RS.Eof
%>
<tr>
<td><%=RS("nombre")%></td>
<td><%=RS("telefono")%></td>
</tr>
<%
RS.MoveNext
Loop

'Cerramos el sistema de conexion
Conn.Close
%>

</table>

<div align="center">
<a href="insertar.html">Añadir un nuevo registro</a><br>
<a href="actualizar1.asp">Actualizar un registro existente</a><br>
<a href="borrar1.asp">Borrar un registro</a><br>
</div>

</BODY>
</HTML>
