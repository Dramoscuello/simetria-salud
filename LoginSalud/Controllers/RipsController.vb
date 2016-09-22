Imports System.Web.Mvc

Namespace Controllers
    Public Class RipsController
        Inherits Controller

        ' GET: Rips
        Function Index() As ActionResult
            Return View()
        End Function
        Dim _Rips As New CodRips
        Public Function ListarControl(ByRef id_u As String) As DataTable
            Return _Rips.ListarControl(id_u)
        End Function

        Public Sub RCargar_Control(ByRef Archi2 As String, ByRef ntba As String, ByRef IdUsuariA As String)
            _Rips.RCargar_Control(Archi2, ntba, IdUsuariA)
        End Sub

        Public Function Excluir(ByRef op As String, ByRef ID As String) As String
            Return _Rips.Excluir(op, ID)
        End Function

        Public Sub Act_dATOSTB(ByRef id As String)
            _Rips.Act_dATOSTB(id)
        End Sub
        Public Sub Act_edades_Q_E_V(ByRef id As String)
            _Rips.Act_edades_Q_E_V(id)
        End Sub
        Public Sub Act_CamposRep(ByRef id As String, ByRef PExcluir As String)
            _Rips.Act_CamposRep(id, PExcluir)
        End Sub
        Public Sub Validar_Consultas(ByRef id As String, ByRef PExcluir As String, ByRef porce As String)
            _Rips.Validar_Consultas(id, PExcluir, porce)
        End Sub
        Public Sub Validar_Hospitalizacion(ByRef id As String, ByRef PExcluir As String)
            _Rips.Validar_Hospitalizacion(id, PExcluir)
        End Sub

        Public Sub Validar_Medicamentos(ByRef id As String, ByRef PExcluir As String)
            _Rips.Validar_Medicamentos(id, PExcluir)
        End Sub
        Public Sub Validar_Nacimientos(ByRef id As String, ByRef PExcluir As String)
            _Rips.Validar_Nacimientos(id, PExcluir)
        End Sub
        Public Sub Validar_Otros_servicios(ByRef id As String, ByRef PExcluir As String, ByRef porce As String)
            _Rips.Validar_Otros_servicios(id, PExcluir, porce)
        End Sub

        Public Sub Validar_Procedimientos(ByRef id As String, ByRef PExcluir As String, ByRef porc As String)
            _Rips.Validar_Procedimientos(id, PExcluir, porc)
        End Sub

        Public Sub Validar_Urgencias(ByRef id As String, ByRef PExcluir As String)
            _Rips.Validar_Urgencias(id, PExcluir)
        End Sub

        Public Sub Validar_Usuarios(ByRef id As String, ByRef PExcluir As String)
            _Rips.Validar_Usuarios(id, PExcluir)
        End Sub
        Public Sub Validar_Transaccion(ByRef id As String, ByRef PExcluir As String)
            _Rips.Validar_Transaccion(id, PExcluir)
        End Sub
        Public Sub TotalFacturado(ByRef IdUsuariA As String, ByRef PExcluir As String)
            _Rips.TotalFacturado(IdUsuariA, PExcluir)
        End Sub
    End Class
End Namespace