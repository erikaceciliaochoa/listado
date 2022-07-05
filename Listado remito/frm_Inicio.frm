VERSION 5.00
Begin VB.Form frm_Inicio 
   Caption         =   "Sistema de Comparación de Programación"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   10590
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu ejemplos 
      Caption         =   "Ejemplos"
      Begin VB.Menu abm_r_n_capas 
         Caption         =   "ABM Remitos"
      End
      Begin VB.Menu ls1 
         Caption         =   "Listados"
      End
      Begin VB.Menu lccc 
         Caption         =   "Listados con corte de control"
      End
      Begin VB.Menu EG 
         Caption         =   "Estadistica - Gráficos"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu edision 
      Caption         =   "Edición"
   End
   Begin VB.Menu formato 
      Caption         =   "Formato"
   End
End
Attribute VB_Name = "frm_Inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub abm_r_n_capas_Click()

frm_remito_n_capas.Show vbModal

End Sub


Private Sub EG_Click()

frm_estadisticas.Show vbModal

End Sub

Private Sub lccc_Click()
frm_listado_remito.Show vbModal
End Sub

Private Sub ls1_Click()
frm_listado.Show vbModal
End Sub

Private Sub salir_Click()
End
End Sub
