VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frm_listado 
   ClientHeight    =   5130
   ClientLeft      =   255
   ClientTop       =   1290
   ClientWidth     =   9285
   Icon            =   "frm_listado.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   9285
   Begin VB.CommandButton Command2 
      Height          =   585
      Left            =   8595
      Picture         =   "frm_listado.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Salir de programa"
      Top             =   4470
      Width           =   585
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3570
      Left            =   90
      TabIndex        =   1
      Top             =   825
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   6297
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Listado de Pedidos"
      TabPicture(0)   =   "frm_listado.frx":06AA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame_fechas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmd_listado_remitos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chk_simple"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chk_entre_fechas"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chk_nombre"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame_nombre"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Listados de Articulos"
      TabPicture(1)   =   "frm_listado.frx":06C6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Frame Frame_nombre 
         Caption         =   "Nombre"
         Height          =   720
         Left            =   3825
         TabIndex        =   14
         Top             =   1260
         Visible         =   0   'False
         Width           =   4410
         Begin VB.TextBox txt_nombre 
            Height          =   300
            Left            =   90
            TabIndex        =   15
            Top             =   270
            Width           =   4245
         End
      End
      Begin VB.CheckBox chk_nombre 
         Caption         =   "Restringir x nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1110
         TabIndex        =   13
         Top             =   1635
         Width           =   2235
      End
      Begin VB.CheckBox chk_entre_fechas 
         Caption         =   "Restringir entre fechas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1110
         TabIndex        =   4
         Top             =   1350
         Width           =   2475
      End
      Begin VB.CheckBox chk_simple 
         Caption         =   "Listado Simple"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1110
         TabIndex        =   3
         Top             =   1080
         Width           =   1650
      End
      Begin VB.CommandButton cmd_listado_remitos 
         Caption         =   "Ejecutar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7485
         TabIndex        =   2
         ToolTipText     =   "Realizar el listado de remitos"
         Top             =   3060
         Width           =   1425
      End
      Begin VB.Frame Frame_fechas 
         Caption         =   "Fechas"
         Height          =   720
         Left            =   3825
         TabIndex        =   8
         Top             =   1260
         Visible         =   0   'False
         Width           =   4410
         Begin VB.TextBox txt_fecha_i 
            Height          =   300
            Left            =   720
            TabIndex        =   12
            Top             =   270
            Width           =   1185
         End
         Begin VB.TextBox txt_fecha_f 
            Height          =   300
            Left            =   2805
            TabIndex        =   11
            Top             =   270
            Width           =   1185
         End
         Begin VB.Label lbl_final 
            AutoSize        =   -1  'True
            Caption         =   "Final"
            Height          =   195
            Left            =   2235
            TabIndex        =   10
            Top             =   330
            Width           =   330
         End
         Begin VB.Label lbl_inicial 
            AutoSize        =   -1  'True
            Caption         =   "Inicial"
            Height          =   195
            Left            =   135
            TabIndex        =   9
            Top             =   323
            Width           =   405
         End
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Listados "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000013&
      Height          =   555
      Left            =   3405
      TabIndex        =   7
      Top             =   75
      Width           =   2130
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Listados "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   555
      Left            =   3390
      TabIndex        =   6
      Top             =   45
      Width           =   2130
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Listados "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3420
      TabIndex        =   5
      Top             =   90
      Width           =   2130
   End
End
Attribute VB_Name = "frm_listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chk_entre_fechas_Click()

Frame_nombre.Visible = False

If chk_entre_fechas.Value = 1 Then
    Frame_fechas.Visible = True
    If txt_fecha_i = "" Then
        txt_fecha_i = "09/10/2006"
    End If
    If txt_fecha_f = "" Then
        txt_fecha_f = "09/10/2006"
    End If
    
Else
    Frame_fechas.Visible = False
End If

End Sub

Private Sub chk_nombre_Click()

Frame_fechas.Visible = False

If chk_nombre.Value = 1 Then
    Frame_nombre.Visible = True
Else
    Frame_nombre.Visible = False
End If

End Sub

Private Sub cmd_listado_remitos_Click()

If chk_simple.Value = 1 Then
    Set Lista_remitos = New Lista_remitos
        
    Lista_remitos.TopMargin = 300
    Lista_remitos.LeftMargin = 300
    Lista_remitos.RightMargin = 300
    
    Set Lista_remitos.DataSource = acceso.Busco_remitos_y_cliente_todos()
    Lista_remitos.Sections("encabezado").Controls("etiqueta10").Caption = "Todos los Remitos"
    Lista_remitos.Show vbModal
End If

If chk_entre_fechas = 1 Then

    If validador(txt_fecha_i, tip_fecha) = False Then
        Exit Sub
    End If
    If validador(txt_fecha_f, tip_fecha) = False Then
        Exit Sub
    End If
    
    Set Lista_remitos = New Lista_remitos
    
    Lista_remitos.LeftMargin = 300
    Lista_remitos.RightMargin = 300
    
    Set Lista_remitos.DataSource = acceso.Busco_remitos_x_fecha(txt_fecha_i, txt_fecha_f)
    Lista_remitos.Sections("encabezado").Controls("etiqueta10").Caption = "Fecha Inicial: " & txt_fecha_i & " Fecha Final: " & txt_fecha_f
    Lista_remitos.Show vbModal
End If

If chk_nombre = 1 Then
    Set Lista_remitos = New Lista_remitos
    
    Lista_remitos.LeftMargin = 300
    Lista_remitos.RightMargin = 300
    
    Set Lista_remitos.DataSource = acceso.Busco_remitos_x_nombre(UCase(txt_nombre))
    Lista_remitos.Sections("encabezado").Controls("etiqueta10").Caption = "Para apellido o Nombres que contengan " & txt_nombre
    Lista_remitos.Show vbModal
End If


End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

' si ud quiere chequear o analizar el final de un formulario puede
' introducir programación en este punto en donde utilizando la variable
' cancel se puede abortar o continuar el proceso de salir del formulario
' cancel = 0 continua con el proceso desalir
' cancel = 1 aborta el proceso de salir y continua dentro del formulario

'If MsgBox("Esta seguro que desea salir del sistema", vbYesNo + vbQuestion, "Atención") = vbYes Then
'Cancel = 0
'Else
'Cancel = 1
'End If

End Sub

