Attribute VB_Name = "UDF"
Public Enum tipos_validacion
    tip_texto
    tip_numero
    tip_fecha
End Enum

'ESTA FUNCION "STRTRAN" TIENE COMO MISION EL PODER CAMBIAR DENTRO
'DE UN STRING UN CARACTER POR OTRO EN TODAS SUS OCURRENCIAS
'ORIGEN DE LA FUNCION : WEB
'UTILIZACION : CAMBIAR COMA DECIMAL POR PUNTO DECIMAL


Public Function StrTran(Cadena As String, Buscar As String, _
Sustituir As String, Optional Veces As Variant) As String

Dim Contador As Integer
Dim Resultado As String
Dim Cambios As Integer
Resultado = ""
Cambios = 0
   
For Contador = 1 To Len(Cadena)
    If Mid(Cadena, Contador, Len(Buscar)) = Buscar Then
        Resultado = Resultado & Sustituir
        If Len(Buscar) > 1 Then
           Contador = Contador + Len(Buscar) - 1
        End If
        ' si se especifica un nº de cambios determinados
        If Not IsMissing(Veces) Then
            Cambios = Cambios + 1
            If Cambios = Veces Then
                Resultado = Resultado & Mid(Cadena, Contador + 1)
                Exit For
            End If
        End If
        
        If Len(Buscar) > 1 Then
            Contador = Contador + Len(Buscar) - 1
        End If
          
    Else
        Resultado = Resultado & Mid(Cadena, Contador, 1)
    End If
Next
   
StrTran = Resultado

End Function
Public Sub carga_combo(ByRef combo As ComboBox, ByRef Record_set As Recordset)

combo.Clear
Do While Record_set.EOF = False
    combo.AddItem Record_set.Fields(1)
    combo.ItemData(combo.NewIndex) = Record_set.Fields(0)
    Record_set.MoveNext
Loop
    
End Sub

Public Sub setea_combo(ByRef combo As ComboBox, ByVal valor As Long)

Dim c As Integer

For c = 0 To combo.ListCount - 1
If combo.ItemData(c) = valor Then
    combo.ListIndex = c
    Exit For
End If
Next

End Sub

Public Sub setea_combo_x_descripcion(ByRef combo As ComboBox, ByVal texto As String)

Dim c As Integer

For c = 0 To combo.ListCount - 1
If UCase(combo.List(c)) = UCase(Trim(texto)) Then
    combo.ListIndex = c
    Exit For
End If
Next

End Sub

Public Function validador(ByVal p_texto As Variant, p_tipo As tipos_validacion) As Boolean

Select Case p_tipo
Case tip_fecha
    If p_texto = "" Then
        MsgBox "La fecha está vacía", vbExclamation, "Importante"
        validador = True
        Exit Function
    End If
    If IsDate(p_texto) Then
        If Not p_texto = Format(CDate(p_texto), "dd/mm/yyyy") Then
            MsgBox "La fecha cargada no es valida", vbCritical
            validador = False
        Else
            validador = True
        End If
    Else
            MsgBox "La fecha no es valida", vbCritical
            validador = False
    End If
Case tip_numero
    If p_texto = "" Then
        MsgBox "El número está vacía", vbExclamation, "Importante"
        validador = True
        Exit Function
    End If
    If Not p_texto = "" Then
    If Not IsNumeric(p_texto) Then
        MsgBox "El número de sucursal es invalido", vbCritical
        validador = False
    Else
        validador = True
    End If
    Else
        validador = False
    End If
Case tip_texto
    If p_texto = "" Then
        MsgBox "El texto esta vacío", vbCritical, "Importante"
        validador = False
    Else
        validador = True
    End If
Case Else
        MsgBox "No configuró el formato de control", vbExclamation, "Importante"
        validador = True
End Select

End Function

