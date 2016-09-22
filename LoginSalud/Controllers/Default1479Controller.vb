Imports System.Web.Mvc
Imports LoginSalud.Models
Namespace Controllers
    Public Class Default1479Controller
        Inherits Controller
        Function Index() As ActionResult
            Return View()
        End Function
        Dim _Resoucion1479 As New Validacion1479
        Public Function ListarErrorDetalle(ByVal IdUsuariA As String) As DataTable
            Return _Resoucion1479.ListarErrorDetalle(IdUsuariA)
        End Function
        Public Function ListarErrorEstructura(ByVal IdUsuariA As String) As DataTable
            Return _Resoucion1479.ListarErrorEstructura(IdUsuariA)
        End Function
        Public Function ListartotalFacturado(ByVal IdUsuariA As String) As DataTable
            Return _Resoucion1479.Total_Facturado(IdUsuariA)
        End Function
        Public Function VerificarID(ByRef temnit As String) As Integer
            Return _Resoucion1479.VerificarID(temnit)
        End Function
        Public Sub Eliminar_registros(ByRef IdUsuariA As String)
            _Resoucion1479.Eliminar_registros(IdUsuariA)
        End Sub
        Public Sub ImportarControl1479(ByRef ruta As String, ByRef IdUsuariA As String)
            _Resoucion1479.ImportarControl1479(ruta, IdUsuariA)
        End Sub
        Public Sub importarD1479(ByRef ruta As String, ByRef IdUsuariA As String)
            _Resoucion1479.importarD1479(ruta, IdUsuariA)
        End Sub
        Public Sub ActualizarEapb(ByRef IdUsuariA As String)
            _Resoucion1479.ActualizarEapb(IdUsuariA)
        End Sub
        Public Sub Validar_Estructura(ByRef IdUsuariA As String)
            _Resoucion1479.Validar_Estructura(IdUsuariA)
        End Sub
        Public Sub Validar_Detalles(ByRef IdUsuariA As String)
            _Resoucion1479.Validar_Detalles(IdUsuariA)
        End Sub

    End Class
End Namespace