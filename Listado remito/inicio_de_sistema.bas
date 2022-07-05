Attribute VB_Name = "inicio_de_sistema"

Public acceso As acceso_remitos


Sub main()

Set acceso = New acceso_remitos
acceso.conexion
frm_Inicio.Show vbModal

End Sub
