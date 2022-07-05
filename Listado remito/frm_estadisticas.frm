VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form frm_estadisticas 
   Caption         =   "Estadísticas de Clientes"
   ClientHeight    =   8355
   ClientLeft      =   465
   ClientTop       =   735
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   11760
   Begin VB.CommandButton cmd_tipo3 
      Caption         =   "Tipo3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3960
      TabIndex        =   3
      Top             =   7800
      Width           =   1485
   End
   Begin VB.CommandButton cmd_tipo2 
      Caption         =   "Tipo 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2280
      TabIndex        =   2
      Top             =   7800
      Width           =   1485
   End
   Begin MSChart20Lib.MSChart g1 
      Height          =   7455
      Left            =   15
      OleObjectBlob   =   "frm_estadisticas.frx":0000
      TabIndex        =   0
      Top             =   0
      Width           =   11535
   End
   Begin VB.CommandButton cmd_tipo1 
      Caption         =   "Tipo 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   600
      TabIndex        =   1
      Top             =   7800
      Width           =   1485
   End
End
Attribute VB_Name = "frm_estadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1

Private Sub cmd_tipo1_Click()

Dim calculo As ADODB.Recordset

Set calculo = acceso.Pedidos_x_clientes()

Dim datos()
ReDim datos(1 To calculo.RecordCount, 1 To 2)
Dim c As Integer


c = 0
calculo.MoveFirst
Do While calculo.EOF = False
    c = c + 1
    datos(c, 1) = calculo!cliente
    datos(c, 2) = calculo!cuanto
    calculo.MoveNext
Loop


g1.ChartData = datos

'g1.Row = 1
'g1.RowLabel = datos(1, 1)

g1.Title = "Cantidad de Pedidos por Clientes"
g1.Title.VtFont.Name = "Arial"
g1.Title.VtFont.Size = 14

g1.ShowLegend = False
g1.AllowSelections = False

g1.RowLabelCount = 3
g1.RowLabelIndex = 2
g1.RowLabel = "Clientes que Mas compraron"
g1.RowLabelIndex = 3
g1.RowLabel = "Año 2007"


g1.chartType = VtChChartType2dBar

End Sub

Private Sub cmd_tipo2_Click()

Dim calculo As ADODB.Recordset

Set calculo = acceso.Pedidos_x_clientes()

Dim datos()
ReDim datos(1 To 2, 1 To calculo.RecordCount)
Dim c As Integer

c = 0
calculo.MoveFirst
Do While Not calculo.EOF
    c = c + 1
    datos(1, c) = calculo!cliente
    datos(2, c) = calculo!cuanto
    calculo.MoveNext
Loop

g1.ChartData = datos

g1.Title = "Cantidad de Pedidos por Clientes"
g1.Title.VtFont.Name = "Arial"
g1.Title.VtFont.Size = 14


g1.AllowSelections = True

g1.ShowLegend = True
g1.Legend.Location.LocationType = VtChLocationTypeBottomRight

g1.chartType = VtChChartType2dBar


End Sub

Private Sub cmd_tipo3_Click()

Dim calculo As ADODB.Recordset

Set calculo = acceso.Pedidos_x_clientes()

Dim datos()
ReDim datos(1 To 2, 1 To calculo.RecordCount)
Dim c As Integer

c = 0
calculo.MoveFirst
Do While Not calculo.EOF
    c = c + 1
    datos(1, c) = calculo!cliente
    datos(2, c) = calculo!cuanto
    calculo.MoveNext
Loop

g1.ChartData = datos

g1.Title = "Cantidad de Pedidos por Clientes"
g1.Title.VtFont.Name = "Arial"
g1.Title.VtFont.Size = 14


g1.Row = 1
'g1.RowLabel = ""

g1.ShowLegend = True
'g1.Legend.Location.LocationType = VtChLocationTypeBottomRight
'g1.Legend.Location.LocationType = VtChLocationTypeLeft
g1.AllowSelections = True
'permite que se puededa seleccionar partes del grafico en tiempo
'de ejecución
g1.Backdrop.Fill.Style = VtFillStyleBrush
g1.CausesValidation = False


g1.chartType = VtChChartType2dPie


End Sub

Private Sub Form_Load()

cmd_tipo1_Click

End Sub
