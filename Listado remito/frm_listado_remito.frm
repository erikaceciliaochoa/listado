VERSION 5.00
Begin VB.Form frm_listado_remito 
   Caption         =   "Listador"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cbo_clientes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2070
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1665
      Width           =   4695
   End
   Begin VB.CommandButton cmd_listado_de_un_cliente 
      Caption         =   "Listado de Pedidos de un Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2070
      TabIndex        =   3
      Top             =   2145
      Width           =   4695
   End
   Begin VB.CommandButton cmd_listado_con_detalle 
      Caption         =   "Listado de Pedidos con Detalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2070
      TabIndex        =   2
      Top             =   1095
      Width           =   4695
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   8070
      Picture         =   "frm_listado_remito.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir de programa"
      Top             =   3105
      Width           =   495
   End
   Begin VB.CommandButton cmd_listado_remitos 
      Caption         =   "Lista de Peidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2070
      TabIndex        =   0
      ToolTipText     =   "Realizar el listado de remitos"
      Top             =   600
      Width           =   4695
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   150
      X2              =   8670
      Y1              =   2985
      Y2              =   2985
   End
   Begin VB.Line Line1 
      X1              =   150
      X2              =   8670
      Y1              =   2865
      Y2              =   2865
   End
End
Attribute VB_Name = "frm_listado_remito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_listado_con_detalle_Click()

Set Rd = New Rd
Rd.DataMember = ""
Set Rd.DataSource = acceso.busco_remitos_y_sus_detalles()
Rd.Show vbModal

End Sub

Private Sub cmd_listado_de_un_cliente_Click()

If cbo_clientes.ListIndex = -1 Then
    MsgBox "No seleccionó ningún cliente"
    Exit Sub
End If

Set Rd = New Rd
Rd.DataMember = ""
Set Rd.DataSource = acceso.busco_remitos_de_un_cliente(cbo_clientes.ItemData(cbo_clientes.ListIndex))
Rd.Show vbModal

End Sub

Private Sub cmd_listado_remitos_Click()

Set Lista_remitos = New Lista_remitos
Set Lista_remitos.DataSource = acceso.Busco_remitos_y_cliente_todos()
Lista_remitos.Show vbModal

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()

carga_combo acceso.leo_tabla_accesoria("t_clientes"), cbo_clientes

End Sub
Private Sub carga_combo(ByRef datos As ADODB.Recordset, ByRef combo As ComboBox)

Do While datos.EOF = False
combo.AddItem datos.Fields(1)
combo.ItemData(combo.NewIndex) = datos.Fields(0)
datos.MoveNext
Loop

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
