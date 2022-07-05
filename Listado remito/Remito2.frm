VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frm_remito_n_capas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orden de Pedidos"
   ClientHeight    =   5895
   ClientLeft      =   3210
   ClientTop       =   1665
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   6390
   Begin VB.TextBox txtsucursal 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Top             =   360
      Width           =   1515
   End
   Begin VB.CommandButton cmdbuscar_cliente 
      Caption         =   ". . ."
      Height          =   315
      Left            =   2025
      TabIndex        =   8
      Top             =   1020
      Width           =   405
   End
   Begin VB.CommandButton cmdagregar 
      Caption         =   "Agregar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4725
      TabIndex        =   16
      Top             =   2565
      Width           =   1485
   End
   Begin VB.TextBox txttotal_detalle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4815
      TabIndex        =   15
      Top             =   2205
      Width           =   1365
   End
   Begin VB.TextBox txtprecio 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2880
      TabIndex        =   14
      Top             =   2205
      Width           =   1335
   End
   Begin VB.TextBox txtcantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1020
      TabIndex        =   13
      Top             =   2220
      Width           =   1065
   End
   Begin VB.TextBox txtn_articulo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1995
      TabIndex        =   12
      Top             =   1845
      Width           =   4185
   End
   Begin VB.TextBox txtid_articulo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   975
      TabIndex        =   10
      Top             =   1845
      Width           =   555
   End
   Begin VB.TextBox txttotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   5070
      TabIndex        =   22
      Top             =   5565
      Width           =   1275
   End
   Begin VB.TextBox txtdescuento 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5070
      TabIndex        =   21
      Top             =   5235
      Width           =   1275
   End
   Begin VB.TextBox txtsubtotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   5070
      TabIndex        =   20
      Top             =   4905
      Width           =   1275
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   600
      Left            =   1920
      Picture         =   "Remito2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Salir del Programa"
      Top             =   5160
      Width           =   600
   End
   Begin VB.CommandButton cmdgrabar 
      Height          =   600
      Left            =   840
      Picture         =   "Remito2.frx":0268
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Graba el registro actual"
      Top             =   5160
      Width           =   600
   End
   Begin VB.CommandButton cmdnuevo 
      Height          =   600
      Left            =   90
      Picture         =   "Remito2.frx":04EA
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Nuevo Registro"
      Top             =   5160
      Width           =   600
   End
   Begin VB.TextBox txtn_cliente 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2445
      TabIndex        =   9
      Top             =   1020
      Width           =   3900
   End
   Begin VB.TextBox txtid_cliente 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      TabIndex        =   7
      Top             =   1020
      Width           =   780
   End
   Begin VB.TextBox txtcomprobante 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      TabIndex        =   6
      Top             =   690
      Width           =   1515
   End
   Begin VB.TextBox txtfecha 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      TabIndex        =   4
      Top             =   30
      Width           =   1515
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalle del Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3405
      Left            =   0
      TabIndex        =   23
      Top             =   1440
      Width           =   6375
      Begin MSFlexGridLib.MSFlexGrid g 
         Height          =   1815
         Left            =   90
         TabIndex        =   31
         Top             =   1515
         Width           =   6240
         _ExtentX        =   11007
         _ExtentY        =   3201
         _Version        =   393216
      End
      Begin VB.CommandButton cmdbuscar_articulo 
         Caption         =   ". . ."
         Height          =   315
         Left            =   1590
         TabIndex        =   11
         Top             =   390
         Width           =   375
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4275
         TabIndex        =   27
         Top             =   780
         Width           =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2205
         TabIndex        =   26
         Top             =   810
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   25
         Top             =   780
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Articulo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   225
         TabIndex        =   24
         Top             =   420
         Width           =   660
      End
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4575
      TabIndex        =   30
      Top             =   5625
      Width           =   450
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Descuento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4095
      TabIndex        =   29
      Top             =   5295
      Width           =   930
   End
   Begin VB.Label laber9 
      AutoSize        =   -1  'True
      Caption         =   "Subtotal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4305
      TabIndex        =   28
      Top             =   4965
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   555
      TabIndex        =   3
      Top             =   1080
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Comprobante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   750
      Width           =   1125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Sucursal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   375
      TabIndex        =   1
      Top             =   450
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   585
      TabIndex        =   0
      Top             =   105
      Width           =   540
   End
End
Attribute VB_Name = "frm_remito_n_capas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset
Dim rs_detalle  As ADODB.Recordset

Dim cf As Byte
Dim nuevo As Boolean

Private Sub Form_Load()

nuevo = False
cf = 1
define_grilla
txtfecha.Text = Date

End Sub

Private Sub cmdagregar_Click()

If validador(Me.txtid_articulo, tip_numero) = False Or Me.txtid_articulo = "" Then
    MsgBox "No es un codigo de articulo valido", vbCritical, "Error Critico"
    txtid_articulo.SetFocus
    Exit Sub
End If
'*************
' controlo que el código del articulo sea valido
' utiliso una función declarada en el área de acceso a datos
' ------------
If acceso.controlo_pk_articulo(txtid_articulo) = False Then
    MsgBox "El código de articulo es inexistente", vbCritical, "Error critico"
    txtid_articulo.SetFocus
    Exit Sub
End If

If validador(Me.txtcantidad, tip_numero) = False Or Me.txtcantidad = "" Then
    txtcantidad.SetFocus
    Exit Sub
End If

If validador(Me.txtprecio, tip_numero) = False Or Me.txtprecio = "" Then
    MsgBox "No es un precio valida", vbCritical, "Error Critico"
    txtprecio.SetFocus
    Exit Sub
End If
'*************
' controlo que el articulo no haya sido cargado ya
'
' ------------
Dim c As Integer
Dim cf1 As Integer

If cf > 1 Then
    For c = 1 To cf - 1
     If g.TextMatrix(c, 5) = txtid_articulo Then
        Exit For
      End If
    Next

    If c <= cf - 1 Then
     MsgBox "Ya cargo este articulo", vbCritical, "Error Critico"
     Exit Sub
    End If
End If
'*************
' agrego nueva fila en el flexGrid
' transfiero los datos ya validados a la grilla
' ------------
 cf = cf + 1
 If cf > g.Rows Then
    g.Rows = cf
 End If
 cf1 = cf - 1
g.TextMatrix(cf1, 1) = txtcantidad.Text
g.TextMatrix(cf1, 2) = txtn_articulo.Text
g.TextMatrix(cf1, 3) = txtprecio.Text
g.TextMatrix(cf1, 5) = txtid_articulo.Text
g.TextMatrix(cf1, 4) = txttotal_detalle.Text

Dim cuenta As Double

For c = 1 To cf1
    cuenta = cuenta + g.TextMatrix(c, 4)
Next

'*************
' calculo el total del la grilla y lo transfiero a sub-total
' calculo valor de total segun operación
' ------------

txtsubtotal.Text = cuenta
txttotal.Text = cuenta - Val(txtdescuento.Text)

cmdagregar.Enabled = False

txtid_articulo.Text = ""
txtid_articulo.SetFocus

End Sub

Private Sub cmdbuscar_articulo_Click()

If txtid_articulo = "" Or validador(Me.txtid_articulo, tip_numero) = False Then
    MsgBox "No existe artículo que buscar", vbCritical, "Importante"
    txtid_articulo.SetFocus
    Exit Sub
End If

Set rs = acceso.busco_articulo(txtid_articulo)

If rs.RecordCount = 0 Then
    MsgBox "No existe ese código para un articulo", vbCritical, "Error Critico"
    txtid_articulo.SetFocus
    txtn_articulo.Text = "Error en código"
    txtprecio.Text = ""
    Exit Sub
End If

rs.MoveFirst

txtn_articulo.Text = rs!n_articulo
txtprecio.Text = rs!precio
txtcantidad.SetFocus

cmdagregar.Enabled = True

txtcantidad.Text = ""
txtcantidad.SetFocus

End Sub

Private Sub cmdbuscar_cliente_Click()

If Me.txtid_cliente = "" Or validador(Me.txtid_cliente, tip_numero) = False Then
    MsgBox "No hay un cliente que buscar", vbCritical, "Importante"
    Me.txtn_cliente.Text = ""
    Me.txtid_cliente.SetFocus
    Exit Sub
End If

Set rs = acceso.busco_cliente(txtid_cliente)

If rs.RecordCount = 0 Then
    MsgBox "No existe ese código para un cliente", vbCritical, "Error Critico"
    txtid_cliente.SetFocus
    txtn_cliente.Text = "Error en el código"
    Exit Sub
End If

txtn_cliente.Text = rs!n_cliente

End Sub

Private Sub cmdgrabar_Click()

Dim respuesta As error_cr

acceso.nro_pedido = txtcomprobante
acceso.nro_sucursal = txtsucursal
acceso.fecha = txtfecha
acceso.id_cliente = txtid_cliente
acceso.total = txttotal
acceso.sub_total = txtsubtotal
acceso.descuento = txtdescuento

respuesta = rn_remitos()

Select Case respuesta
Case fecha
    txtfecha.SetFocus
    Exit Sub
Case id_cliente
    txtid_articulo.SetFocus
    Exit Sub
Case nro_pedido
    txtcomprobante.SetFocus
    Exit Sub
Case nro_sucursal
    txtsucursal.SetFocus
    Exit Sub
Case sub_total
    txtsubtotal.SetFocus
    Exit Sub
Case total
    txttotal.SetFocus
    Exit Sub
Case Ok

End Select

If acceso.control_PK_cliente(acceso.id_cliente) = False Then
    MsgBox "El código de cliente no existe", vbCritical, "Importante"
    Exit Sub
End If

Dim status As Boolean

If acceso.control_PK_remitos(acceso.nro_pedido, acceso.nro_sucursal) = False Then
' agregar registro
 If acceso.Inserta_Cabecera_Remitos = False Then
    MsgBox "Fracaso la inserción de la cabecera del remito", vbInformation, "Importante"
    Exit Sub
 End If
 
Else
    MsgBox "Este remito ya esta grabado", vbCritical, "Importante"
    Exit Sub
End If

Dim c As Byte

For c = 1 To cf - 1
    acceso.cantidad = g.TextMatrix(c, 1)
    acceso.precio = g.TextMatrix(c, 3)
    acceso.id_articulo = g.TextMatrix(c, 5)
    If acceso.Inserta_Detalle_Remito() = False Then
        MsgBox "Fracaso la inserción del detalle del remito", vbInformation, "Importante"
        acceso.Borrar_cabecera_Remito
        acceso.Borrar_detalle_Remito
        Exit Sub
    End If

Next

MsgBox "Se grabo con exito", vbInformation, "Mensaje del sistema"

cmdgrabar.Enabled = False

End Sub

Private Sub cmdnuevo_Click()

nuevo = True
txtfecha.Text = Date
g.Rows = 1
cf = 1
blanquear
cmdgrabar.Enabled = True
txtfecha.SetFocus

End Sub
Private Sub blanquear()

Dim objtxt As Control

For Each objtxt In Me.Controls
If TypeOf objtxt Is TextBox Then
    objtxt = ""
End If
Next

End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub define_grilla()

g.Rows = 1
g.Cols = 6

g.AllowUserResizing = flexResizeColumns
g.ColWidth(0) = 220
g.ColWidth(2) = 2700

g.TextMatrix(0, 1) = "Cantidad"
g.TextMatrix(0, 2) = "Item"
g.TextMatrix(0, 3) = "Monto"
g.TextMatrix(0, 4) = "Total"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

' si ud quiere chequear o analizar el final de un formulario puede
' introducir programación en este punto en donde utilizando la variable
' cancel se puede abortar o continuar el proceso de salir del formulario
' cancel = 0 continua con el proceso desalir
' cancel = 1 aborta el proceso de salir y continua dentro del formulario

If MsgBox("Esta seguro que desea salir del sistema", vbYesNo + vbQuestion, "Atención") = vbYes Then
Cancel = 0
Else
Cancel = 1
End If

End Sub


Private Sub txtcantidad_Change()

If IsNumeric(txtcantidad) Then
    txttotal_detalle.Text = Val(txtprecio.Text) * Val(txtcantidad.Text)
End If

End Sub

Private Sub txtcantidad_LostFocus()
If validador(Me.txtcantidad, tip_numero) = False Then
    Me.txtcantidad.SetFocus
End If

End Sub

Private Sub txtcomprobante_LostFocus()

If validador(Me.txtcomprobante, tip_numero) = False Then
    Me.txtcomprobante.SetFocus
End If

End Sub

Private Sub txtdescuento_Change()

Dim calculo As Double

calculo = txtsubtotal.Text
txttotal.Text = calculo - txtdescuento.Text

End Sub


Private Sub txtfecha_LostFocus()

If validador(Me.txtfecha, tip_fecha) = False Then
    Me.txtfecha.SetFocus
End If

End Sub

Private Sub txtid_articulo_LostFocus()
If validador(txtid_articulo, tip_numero) = False Then
    Me.txtid_articulo.SetFocus
End If
End Sub

Private Sub txtid_cliente_LostFocus()
If validador(Me.txtid_cliente, tip_numero) = False Then
    Me.txtid_cliente.SetFocus
End If
End Sub

Private Sub txtprecio_Change()
    If IsNumeric(txtprecio) = True Then
    txttotal_detalle.Text = Val(txtprecio.Text) * Val(txtcantidad.Text)
    End If
End Sub

Private Sub txtprecio_LostFocus()
If validador(Me.txtprecio, tip_numero) = False Then
    Me.txtprecio.SetFocus
End If
End Sub


Private Sub txtsucursal_LostFocus()

If validador(Me.txtsucursal, tip_numero) = False Then
    Me.txtsucursal.SetFocus
End If

End Sub

