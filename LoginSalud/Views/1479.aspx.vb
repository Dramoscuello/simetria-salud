Imports System.IO
Imports ClosedXML.Excel
Imports LoginSalud.Controllers

Public Class _1479
    Inherits System.Web.UI.Page
    Dim idusu As String
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Response.Cache.SetCacheability(HttpCacheability.ServerAndNoCache)
        Response.Cache.SetAllowResponseInBrowserHistory(False)
        Response.Cache.SetNoStore()
        If Session("usuario") IsNot Nothing Then
            idusu = Session("usuario")
        Else
            Response.Redirect("~/Ingreso")
        End If
        ButtonInforme.Enabled = False
    End Sub
    Dim tipoIdent, temnit As String
    Dim nombreArchivo As String
    Private Function Contar_Columnas(ByRef ArchiDir As String) As Boolean
        Dim Archidir3 As String
        Dim K As Integer = 0 : Dim Lat As Integer = 0
        Dim resul As Boolean
        Archidir3 = Replace(ArchiDir, "\", "/")
        Dim sr_ As New StreamReader(Archidir3)
        Dim Array_Consultas() As String
        Dim LINEA() As String
        Try
            K = 0
            LINEA = sr_.ReadLine.Split(",")
            Lat = (File.ReadAllLines(Archidir3).Count - 1)
            If Lat <> LINEA(5) Then
                MsgBox("El Numero de Registros No Conciden con la cantidad de registros (Linea de control)", MsgBoxStyle.Exclamation, "CONTAR COLUMNAS")
                Return False
            End If
            Do While (K < LINEA(5))
                K = K + 1
                Array_Consultas = sr_.ReadLine.Split(",")
                If Array_Consultas.Length <> 25 Then
                    MsgBox("Error en la Linea " & K & " El Numero de campos No Conciden", MsgBoxStyle.Exclamation, "CONTAR COLUMNAS")
                    Return False
                Else
                    Return True
                End If
            Loop
            sr_.Close()
            'Return resul
        Catch ex As Exception
            Return Nothing
        Finally
            sr_.Close()
        End Try
    End Function
    Private Sub EmpezarValidacion(ByRef ruta As String)
        Dim InicioV As New Default1479Controller
        Try
            temnit = Mid(nombreArchivo, 21, 12)
            If tipoIdent = "DI" Or tipoIdent = "DE" Then
                If InicioV.VerificarID(temnit) = 0 Then
                    ClientScript.RegisterStartupScript(Me.GetType, "alerta", "swal('¡Alerta!','¡El Numero De Identificación No Corresponde A La Entidad Reportadora!', 'warning')", True)
                    Exit Sub
                Else
                    If Contar_Columnas(ruta) = True Then
                        InicioV.Eliminar_registros(idusu)
                        InicioV.ImportarControl1479(ruta, idusu)
                        InicioV.importarD1479(ruta, idusu)
                        InicioV.ActualizarEapb(idusu)
                        InicioV.Validar_Estructura(idusu)
                        InicioV.Validar_Detalles(idusu)
                        ButtonInforme.Enabled = True
                        'Thread.Sleep(2000)
                        ClientScript.RegisterStartupScript(Me.GetType, "ok", "swal('¡Validación terminada!','¡descargue el informe!', 'success')", True)

                    End If
                End If
            Else
                ClientScript.RegisterStartupScript(Me.GetType, "alerta", "swal('¡Alerta!','¡ Tipo de identificacion la entidad reportadora invalido !', 'warning')", True)
                Exit Sub
            End If
        Catch ex As Exception
        End Try
    End Sub


    Private Sub Iniciar()
        nombreArchivo = FileUploadImportar.FileName
        Dim largo1 As Integer = 0
        largo1 = Len(My.Computer.FileSystem.GetName(nombreArchivo))
        If largo1 < 36 Or largo1 > 36 Then
            ClientScript.RegisterStartupScript(Me.GetType, "alerta", "swal('¡Alerta!','¡Error en Nombre de Archivo (longitud del archivo ) !', 'warning')", True)
            Exit Sub
        End If
        Dim VRMODULO As String
        VRMODULO = LSet(nombreArchivo, 10)
        If VRMODULO <> "TEC120NPOS" Then
            ClientScript.RegisterStartupScript(Me.GetType, "alerta", "swal('¡Alerta!','¡El nombre del archivo  no corresponde con la definición del anexo (TEC120NPOS) !', 'warning')", True)
            Exit Sub
        End If
        Try
            If IO.Directory.Exists(Server.MapPath(nombreArchivo & idusu & "/")) Then
                For Each item As String In Directory.GetFiles(Server.MapPath(nombreArchivo & idusu & "/"))
                    File.Delete(item)
                Next
            End If

            My.Computer.FileSystem.CreateDirectory(Server.MapPath(nombreArchivo & idusu & "/"))
            tipoIdent = Mid(nombreArchivo, 19, 2)
            Dim x As String = Server.MapPath(nombreArchivo & idusu & "/")
            Dim path As String = FileUploadImportar.PostedFile.FileName
            Dim source As String = Replace(x, "\", "/")
            If Not String.IsNullOrEmpty(path) Then
                Dim ImageFiles As HttpFileCollection = Request.Files
                For i As Integer = 0 To ImageFiles.Count - 1
                    Dim file As HttpPostedFile = ImageFiles(i)
                    file.SaveAs(Server.MapPath(nombreArchivo & idusu & "/") & file.FileName)
                Next
                EmpezarValidacion(source + nombreArchivo)
                My.Computer.FileSystem.DeleteFile(Server.MapPath(nombreArchivo & idusu & "/") + path)
                path = Nothing
            ElseIf String.IsNullOrEmpty(path) Then
                ClientScript.RegisterStartupScript(Me.GetType, "alerta", "swal('¡Alerta!','Debe seleccionar el archivo de 1479', 'warning')", True)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Protected Sub ButtonInforme_Click(sender As Object, e As EventArgs) Handles ButtonInforme.Click
        Dim Errores As New Default1479Controller
        Try
            Dim wb As New XLWorkbook()
            Dim Detalle, Estructura, Total_FAc, ok_ As New DataTable
            Detalle = Errores.ListarErrorDetalle(idusu)
            Estructura = Errores.ListarErrorEstructura(idusu)
            Total_FAc = Errores.ListartotalFacturado(idusu)
            If Detalle.Rows.Count > 0 Then
                wb.Worksheets.Add(Detalle, "Errores_Detalle")
            End If

            If Estructura.Rows.Count > 0 Then
                wb.Worksheets.Add(Estructura, "Errores_Estructura")
            End If
            If Total_FAc.Rows.Count > 0 Then
                wb.Worksheets.Add(Total_FAc, "Total Facturado")
            End If
            ok_.Columns.Add(" ")
            ok_.Rows.Add(" ")
            wb.Worksheets.Add(ok_, ".")
            wb.Author = "Simetria"
            Response.Clear()
            Response.Buffer = True
            Response.Charset = ""
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            Response.AddHeader("content-disposition", "attachment;filename=Validacion_1479.xlsx")
            Dim MyMemoryStream As New MemoryStream()
            wb.SaveAs(MyMemoryStream)
            MyMemoryStream.WriteTo(Response.OutputStream)
            Response.Flush()
            Response.End()
        Catch ex As Exception
            If ex.InnerException Is Nothing Then
                ClientScript.RegisterStartupScript(Me.GetType, "error1", "swal('¡Error!'," + ex.Message.ToString() + ", 'error')", True)
            Else
                ClientScript.RegisterStartupScript(Me.GetType, "error2", "swal('¡Error!'," + ex.InnerException.Message.ToString() + ", 'error')", True)
            End If
        End Try
    End Sub

    Protected Sub ButtonValidar_Click(sender As Object, e As EventArgs) Handles ButtonValidar.Click
        Iniciar()
    End Sub
End Class