<TITLE>Actualizar2.asp</TITLE>
</HEAD>
<BODY>

<%
'Recogemos los valores del formulario
nombre=Request.Form("nombre")
telefono= Request.Form("telefono")

'Instanciamos y abrimos nuestro objeto conexion 
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open "Mibase"

'Ahora creamos la sentencia SQL 
sSQL="Update Clientes Set telefono='" & telefono & "' Where nombre='" & nombre & "'"

'Ejecutamos la orden 
set RS = Conn.Execute(sSQL)
%>

<h1><div align="center">Registro Actualizado</div></h1>
<div align="center"><a href="lectura.asp">Visualizar el contenido de la base</a></div>

<%
'Cerramos el sistema de conexion
Conn.Close
%>

</BODY>
</HTML>
