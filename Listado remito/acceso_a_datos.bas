Attribute VB_Name = "acceso_a_dato"
Public Type registro_cr
    nro_pedido As String
    nro_sucursal As String
    id_cliente As String
    fecha As String
    sub_total As String
    descuento As String
    total As String
End Type

Public Type registro_dr
    id_articulo As String
    cantidad As String
    precio As String
End Type

'Public cn As ADODB.Connection
    
Public Sub cierra_conexion(ByVal cn As ADODB.Connection)
    cn.Close
End Sub


Public Function conexion() As ADODB.Connection

Dim ubi As String

'le asigna a ubi el path de ubicación del sistema
ubi = App.Path

Set cn = New ADODB.Connection
cn.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ubi & "\remito.mdb;Persist Security Info=False")
cn.CursorLocation = adUseClient
Set conexion = cn

End Function

Public Function leo_tabla_accesoria(ByRef cn As ADODB.Connection, ByVal tabla As String) As ADODB.Recordset

'declara el nuevo objeto recordSet que utlizará para cargar los combos
Dim rsa1 As ADODB.Recordset
Set rsa1 = New Recordset
'dimensiona una variable string para crear el SQL conque cargara el recordSet
Dim sql As String
'crea el SQL
sql = "select * from " & Trim(tabla) & " order by 2"
'abra la consulta con el recurso SQL, con la conexión cn, con el tipo de
'vinculación adOpenForwardOnly, y de solo lectura
rsa1.Open sql, cn, adOpenForwardOnly, adLockReadOnly

Set leo_tabla_accesoria = rsa1

End Function

Public Function control_PK_remitos(ByRef cn As ADODB.Connection, ByVal p_nro_pedido As Long, ByVal p_nro_sucursal As Integer) As Boolean

Dim sql As String

sql = "select * from t_pedidos "
sql = sql & "where nro_pedido = " & p_nro_pedido
sql = sql & " And nro_sucursal = " & p_nro_sucursal

Set consulta = New ADODB.Command
consulta.CommandText = sql
consulta.CommandType = adCmdText
consulta.ActiveConnection = cn

Set rs = consulta.Execute

If rs.RecordCount = 1 Then
    control_PK_remitos = True
Else
    control_PK_remitos = False
End If

End Function

Public Function control_PK_cliente(ByRef cn As ADODB.Connection, ByVal p_id_cliente As Long) As Boolean

sql = "select * from t_clientes "
sql = sql & "where id_cliente = " & p_id_cliente

Set consulta = New ADODB.Command
Set rs = New ADODB.Recordset

consulta.CommandText = sql
consulta.CommandType = adCmdText
consulta.ActiveConnection = cn

Set rs = consulta.Execute

If rs.RecordCount = 1 Then
    control_PK_cliente = True
Else
    control_PK_cliente = False
End If

End Function

Public Function Inserta_Cabecera_Remitos(ByRef cn As ADODB.Connection, rc As registro_cr) As Boolean

Dim x_fecha As String

x_fecha = "#" & Format(rc.fecha, "mm/dd/yyyy") & "#"
rc.descuento = IIf(rc.descuento = "", "0", "0.00")
rc.total = StrTran(rc.total, ",", ".")
rc.sub_total = StrTran(rc.sub_total, ",", ".")

i_txt = "insert into t_pedidos (nro_pedido, nro_sucursal, "
i_txt = i_txt & "id_cliente, fecha, sub_total, descuentos, "
i_txt = i_txt & "total) values ( " & rc.nro_pedido & ", "
i_txt = i_txt & rc.nro_sucursal & ", " & rc.id_cliente & ", "
i_txt = i_txt & x_fecha & ", " & rc.sub_total & ", "
i_txt = i_txt & rc.descuento & ", " & rc.total & ")"

Set consulta = New ADODB.Command
consulta.ActiveConnection = cn
consulta.CommandText = i_txt
consulta.CommandType = adCmdText
consulta.Execute

Inserta_Cabecera_Remitos = True

End Function

Public Function Inserta_Detalle_Remito(ByRef cn As ADODB.Connection, _
ByVal p_nro_pedido As String, ByVal p_nro_sucursal As String, _
ByRef datos() As registro_dr) As Boolean

Dim c As Long, d As Long

d = UBound(datos, 1) - 1

Set consulta = New ADODB.Command
consulta.ActiveConnection = cn

consulta.CommandType = adCmdText

Dim a As String, b As String


For c = 1 To d

a = StrTran(datos(c).cantidad, ",", ".")
b = StrTran(datos(c).precio, ",", ".")

i_txt = "insert into t_detalles_pedidos values (" & p_nro_pedido & ", "
i_txt = i_txt & p_nro_sucursal & ", " & datos(c).id_articulo & ", "
i_txt = i_txt & a & ", "
i_txt = i_txt & b & ")"

consulta.CommandText = i_txt
consulta.Execute

Next

Inserta_Detalle_Remito = True

End Function

Public Sub Borrar_cabecera_Remito(ByRef cn As ADODB.Connection, _
p_nro_pedido As String, p_nro_sucursal As String)

' esta programación falfa realizarla

End Sub

Public Function controlo_pk_articulo(cn As ADODB.Connection, p_id_articulo As String) As Boolean

Dim sql_txt As String
sql_txt = "select id_articulo, n_articulo, precio from t_articulos where id_articulo = " & p_id_articulo

Set consulta = New ADODB.Command
consulta.CommandText = sql_txt
consulta.CommandType = adCmdText
consulta.ActiveConnection = cn

Set rs = consulta.Execute

If rs.RecordCount = 0 Then
    controlo_pk_articulo = False
Else
    controlo_pk_articulo = True
End If

End Function

Public Function busco_cliente(ByRef cn As ADODB.Connection, ByVal p_id_cliente As String) As ADODB.Recordset

Dim sql_txt As String
sql_txt = "select id_cliente, n_cliente from t_clientes where id_cliente = " & p_id_cliente

Set consulta = New ADODB.Command
consulta.CommandText = sql_txt
consulta.CommandType = adCmdText
consulta.ActiveConnection = cn

Set busco_cliente = consulta.Execute

End Function
Public Function busco_articulo(ByRef cn As ADODB.Connection, ByVal p_id_articulo As String) As ADODB.Recordset

Dim sql_txt As String
sql_txt = "select id_articulo, n_articulo, precio from t_articulos where id_articulo = " & p_id_articulo

Set consulta = New ADODB.Command
consulta.CommandText = sql_txt
consulta.CommandType = adCmdText
consulta.ActiveConnection = cn

Set busco_articulo = consulta.Execute

End Function
Public Function Busco_remitos_y_cliente_todos(ByRef cn As ADODB.Connection) As ADODB.Recordset

Dim sql As String
Dim consulta As ADODB.Command

sql = "select p.nro_pedido as pedido, p.fecha, c.n_cliente as cliente " & _
      "from t_clientes c, t_pedidos p " & _
      "where c.id_cliente = p.id_cliente " & _
      "order by p.fecha, p.nro_pedido"

Set consulta = New ADODB.Command
consulta.ActiveConnection = cn
consulta.CommandType = adCmdText
consulta.CommandText = sql

Set Busco_remitos_y_cliente_todos = consulta.Execute

End Function
Public Function busco_remitos_de_un_cliente(ByVal id As Integer) As ADODB.Recordset

Dim conexion As New ADODB.Connection
conexion.Provider = "MSDataShape"

Dim ubi As String
ubi = App.Path
'la conexion con la base de datos no puede ser via el proveedor
'tipico de ACCESS pues no
'permite la creación de estructuras de datos devueltos en recordset
'por ello se crea un acceso a datos via ODBC,
'se debe crear una conexion a la base de datos en el acceso a datos del
'ODBC

conexion.Open "Shape Provider=MSDASQL;dsn=clase10"

Dim sql As String
sql = "SHAPE {select tp.*, tc.n_cliente" _
+ " from t_pedidos tp, t_clientes tc" _
+ " where tp.id_cliente = tc.id_cliente" _
+ " and tp.id_cliente = " & id & "} AS pedido" _
+ " APPEND ({select dp.*, ta.n_articulo," _
+ " dp.precio*dp.cantidad as subtotal" _
+ " from t_detalles_pedidos dp, t_articulos ta" _
+ " where dp.id_articulo = ta.id_articulo} AS detalle" _
+ " RELATE nro_pedido TO nro_pedido, nro_sucursal TO nro_sucursal)"
    
Dim miRs As New ADODB.Recordset
miRs.Open sql, conexion
 
Set busco_remitos_de_un_cliente = miRs

'conexion.Close

End Function
Public Function busco_remitos_y_sus_detalles() As ADODB.Recordset

Dim conexion As New ADODB.Connection
conexion.Provider = "MSDataShape"
Dim ubi As String
ubi = App.Path
'la conexion con la base de datos no puede ser via el proveedor
'tipico de access pues no
'permite la creación de estructuras de datos devueltos en recordset
'por ello se crea un acceso a datos a travez de ODBC, se debe crear una
'conexion a la base
conexion.Open "Shape Provider=MSDASQL;dsn=clase10"

Dim sql As String
sql = "SHAPE {select tp.*, tc.n_cliente" _
+ " from  t_pedidos tp, t_clientes tc" _
+ " where tp.id_cliente = tc.id_cliente} AS pedido" _
+ " APPEND ({select dp.*, ta.n_articulo," _
+ " dp.precio*dp.cantidad as subtotal " _
+ " from t_detalles_pedidos dp, t_articulos ta" _
+ " where dp.id_articulo = ta.id_articulo} AS detalle" _
+ " RELATE nro_pedido TO nro_pedido, nro_sucursal TO nro_sucursal)"
    
Dim miRs As New ADODB.Recordset
miRs.Open sql, conexion
 
Set busco_remitos_y_sus_detalles = miRs

'conexion.Close

End Function

Public Function Pedidos_x_clientes(ByRef cn As ADODB.Connection) As ADODB.Recordset

Dim sql As String
Dim consulta As ADODB.Command

sql = "select c.n_cliente as cliente, count(*) as cuanto " & _
      "from t_clientes c, t_pedidos p " & _
      "where c.id_cliente = p.id_cliente " & _
      " group by c.n_cliente, c.id_cliente "

Set consulta = New ADODB.Command
consulta.ActiveConnection = cn
consulta.CommandType = adCmdText
consulta.CommandText = sql

Set Pedidos_x_clientes = consulta.Execute

End Function


