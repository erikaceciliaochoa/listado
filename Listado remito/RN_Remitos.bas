Attribute VB_Name = "Regla_de_negocios"
Public Enum error_cr
    nro_pedido
    nro_sucursal
    id_cliente
    fecha
    sub_total
    descuento
    total
    Ok
End Enum

Public Function rn_remitos() As error_cr
If IsDate(acceso.fecha) Then
    If Not acceso.fecha = Format(CDate(acceso.fecha), "dd/mm/yyyy") Then
        MsgBox "La fecha cargada no es valida", vbCritical
        rn_remitos = fecha
        Exit Function
    End If
Else
        MsgBox "La fecha no es valida", vbCritical
        rn_remitos = fecha
        Exit Function
End If

If Not acceso.nro_sucursal = "" Then
If Not IsNumeric(acceso.nro_sucursal) Then
    MsgBox "El numero de sucursal es invalido", vbCritical
    rn_remitos = nro_sucursal
    Exit Function
End If
Else
    rn_remitos = nro_sucursal
    Exit Function
End If

If Not acceso.nro_pedido = "" Then
If Not IsNumeric(acceso.nro_pedido) Then
    MsgBox "El numero de comprobante es invalido", vbCritical
    rn_remitos = nro_pedido
    Exit Function
End If
Else
    rn_remitos = nro_pedido
    Exit Function
End If

If Not acceso.sub_total = "" Then
If Not IsNumeric(acceso.sub_total) Then
    MsgBox "El subtotal  no es numerico", vbCritical
    rn_remitos = sub_total
    Exit Function
End If
Else
    rn_remitos = sub_total
    Exit Function
End If

If acceso.descuento <> "" Then
If Not IsNumeric(acceso.descuento) Then
    MsgBox "El descuento  no es numerico", vbCritical
    rn_remitos = descuento
     Exit Function
End If
End If

If Not acceso.total = "" Then
If Not IsNumeric(acceso.total) Then
    MsgBox "El total  no es numerico", vbCritical
    rn_remitos = total
    Exit Function
End If
Else
    rn_remitos = total
    Exit Function
End If

If Not acceso.id_cliente = "" Then
If Not IsNumeric(acceso.id_cliente) Then
    MsgBox "El codigo de cliente  no es numerico", vbCritical
    rn_remitos = id_cliente
    Exit Function
End If
Else
    rn_remitos = id_cliente
    Exit Function
End If

rn_remitos = Ok

End Function
