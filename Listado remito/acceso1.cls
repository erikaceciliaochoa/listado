VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "acceso_remitos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public cn As ADODB.Connection

Public nro_pedido As String
Public nro_sucursal As String
Public id_cliente As String
Public fecha As String
Public sub_total As String
Public descuento As String
Public total As String


Public id_articulo As String
Public cantidad As String
Public precio As String
    
Public Sub cierra_conexion()
    Me.cn.Close
End Sub


Public Sub conexion() 'As ADODB.Connection

Dim ubi As String

'le asigna a ubi el path de ubicación del sistema
ubi = App.Path

Set Me.cn = New ADODB.Connection
Me.cn.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ubi & "\remito.mdb;Persist Security Info=False")
Me.cn.CursorLocation = adUseClient


End Sub

Public Function leo_tabla_accesoria(ByVal tabla As String) As ADODB.Recordset

'declara el nuevo objeto recordSet que utlizará para cargar los combos
Dim rsa1 As ADODB.Recordset
Set rsa1 = New Recordset
'dimensiona una variable string para crear el SQL conque cargara el recordSet
Dim sql As String
'crea el SQL
sql = "select * from " & Trim(tabla) & " order by 2"
'abra la consulta con el recurso SQL, con la conexión cn, con el tipo de
'vinculación adOpenForwardOnly, y de solo lectura
rsa1.Open sql, Me.cn, adOpenForwardOnly, adLockReadOnly

Set leo_tabla_accesoria = rsa1

End Function

Public Function control_PK_remitos(ByVal p_nro_pedido As Long, ByVal p_nro_sucursal As Integer) As Boolean

Dim sql As String

sql = "select * from t_pedidos "
sql = sql & "where nro_pedido = " & p_nro_pedido
sql = sql & " And nro_sucursal = " & p_nro_sucursal

Set consulta = New ADODB.Command
consulta.CommandText = sql
consulta.CommandType = adCmdText
consulta.ActiveConnection = Me.cn

Set rs = consulta.Execute

If rs.RecordCount = 1 Then
    control_PK_remitos = True
Else
    control_PK_remitos = False
End If

End Function

Public Function control_PK_cliente(ByVal p_id_cliente As Long) As Boolean

sql = "select * from t_clientes "
sql = sql & "where id_cliente = " & p_id_cliente

Set consulta = New ADODB.Command
consulta.CommandText = sql
consulta.CommandType = adCmdText
consulta.ActiveConnection = Me.cn

Set rs = consulta.Execute

If rs.RecordCount = 1 Then
    control_PK_cliente = True
Else
    control_PK_cliente = False
End If

End Function

Public Function Inserta_Cabecera_Remitos() As Boolean

Dim x_fecha As String

x_fecha = "#" & Format(Me.fecha, "mm/dd/yyyy") & "#"

Me.total = StrTran(Me.total, ",", ".")
Me.sub_total = StrTran(Me.sub_total, ",", ".")
Me.descuento = StrTran(Me.descuento, ",", ".")
Me.descuento = IIf(Me.descuento = "", "0", Me.descuento)

i_txt = "insert into t_pedidos (nro_pedido, nro_sucursal, "
i_txt = i_txt & "id_cliente, fecha, sub_total, descuentos, "
i_txt = i_txt & "total) values ( " & Me.nro_pedido & ", "
i_txt = i_txt & Me.nro_sucursal & ", " & Me.id_cliente & ", "
i_txt = i_txt & x_fecha & ", " & Me.sub_total & ", "
i_txt = i_txt & Me.descuento & ", " & Me.total & ")"

Set consulta = New ADODB.Command
consulta.ActiveConnection = Me.cn
consulta.CommandText = i_txt
consulta.CommandType = adCmdText
consulta.Execute

Inserta_Cabecera_Remitos = True

End Function

Public Function Inserta_Detalle_Remito() As Boolean

Set consulta = New ADODB.Command
consulta.ActiveConnection = Me.cn
consulta.CommandType = adCmdText

Me.cantidad = StrTran(Me.cantidad, ",", ".")
Me.precio = StrTran(Me.precio, ",", ".")

i_txt = "insert into t_detalles_pedidos values (" & Me.nro_pedido & ", "
i_txt = i_txt & Me.nro_sucursal & ", " & Me.id_articulo & ", "
i_txt = i_txt & Me.cantidad & ", "
i_txt = i_txt & Me.precio & ")"

consulta.CommandText = i_txt
consulta.Execute

Inserta_Detalle_Remito = True

End Function

Public Sub Borrar_cabecera_Remito()

' esta programación falfa realizarla

End Sub
Public Sub Borrar_detalle_Remito()

' esta programación falfa realizarla

End Sub


Public Function controlo_pk_articulo(p_id_articulo As String) As Boolean

Dim sql_txt As String
sql_txt = "select id_articulo, n_articulo, precio from t_articulos where id_articulo = " & p_id_articulo

Set consulta = New ADODB.Command
consulta.CommandText = sql_txt
consulta.CommandType = adCmdText
consulta.ActiveConnection = Me.cn

Set rs = consulta.Execute

If rs.RecordCount = 0 Then
    controlo_pk_articulo = False
Else
    controlo_pk_articulo = True
End If

End Function
Public Function busco_cliente(ByVal p_id_cliente As String) As ADODB.Recordset

Dim sql_txt As String
sql_txt = "select id_cliente, n_cliente from t_clientes where id_cliente = " & p_id_cliente

Set consulta = New ADODB.Command
consulta.CommandText = sql_txt
consulta.CommandType = adCmdText
consulta.ActiveConnection = Me.cn

Set busco_cliente = consulta.Execute

End Function
Public Function busco_articulo(ByVal p_id_articulo As String) As ADODB.Recordset

Dim sql_txt As String
sql_txt = "select id_articulo, n_articulo, precio from t_articulos where id_articulo = " & p_id_articulo

Set consulta = New ADODB.Command
consulta.CommandText = sql_txt
consulta.CommandType = adCmdText
consulta.ActiveConnection = Me.cn

Set busco_articulo = consulta.Execute

End Function
Public Function Busco_remitos_y_cliente_todos() As ADODB.Recordset

Dim sql As String
Dim consulta As ADODB.Command

sql = "select p.nro_pedido as pedido, p.fecha, c.n_cliente as cliente, p.total " & _
      "from t_clientes c, t_pedidos p " & _
      "where c.id_cliente = p.id_cliente " & _
      "order by p.fecha, p.nro_pedido"

Set consulta = New ADODB.Command
consulta.ActiveConnection = Me.cn
consulta.CommandType = adCmdText
consulta.CommandText = sql

Set Busco_remitos_y_cliente_todos = consulta.Execute

End Function
Public Function Busco_remitos_x_fecha(fecha_i As String, fecha_f As String) As ADODB.Recordset

Dim sql As String
Dim consulta As ADODB.Command

fecha_i = "#" & Format(fecha_i, "yyyy/mm/dd") & "#"
fecha_f = "#" & Format(fecha_f, "yyyy/mm/dd") & "#"


sql = "select p.nro_pedido as pedido, p.fecha, c.n_cliente as cliente, p.total" & _
      " from t_clientes c, t_pedidos p" & _
      " where c.id_cliente = p.id_cliente" & _
      " and p.fecha >= " & fecha_i & _
      " and p.fecha <= " & fecha_f & _
      " order by p.fecha, p.nro_pedido"

Set consulta = New ADODB.Command
consulta.ActiveConnection = Me.cn
consulta.CommandType = adCmdText
consulta.CommandText = sql

Set Busco_remitos_x_fecha = consulta.Execute

End Function
Public Function Busco_remitos_x_nombre(texto As String) As ADODB.Recordset

Dim sql As String
Dim consulta As ADODB.Command

sql = "select p.nro_pedido as pedido, p.fecha, c.n_cliente as cliente, p.total " & _
      " from t_clientes c, t_pedidos p" & _
      " where c.id_cliente = p.id_cliente" & _
      " and c.n_cliente like '%" & Trim(texto) & "%' " & _
      " order by p.fecha, p.nro_pedido"

Set consulta = New ADODB.Command
consulta.ActiveConnection = Me.cn
consulta.CommandType = adCmdText
consulta.CommandText = sql

Set Busco_remitos_x_nombre = consulta.Execute

End Function
Public Function Pedidos_x_clientes() As ADODB.Recordset

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
'tipico de access (OLEDB provider) pues no
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






