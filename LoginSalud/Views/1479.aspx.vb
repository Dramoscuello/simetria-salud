Imports System.IO

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
        Try
            Dim Archidir3 As String
            Dim K As Integer = 0 : Dim Lat As Integer = 0
            Dim resul As Boolean
            Archidir3 = Replace(ArchiDir, "\", "/")
            Dim sr As New StreamReader(Archidir3)
            Dim Array_Consultas() As String
            Dim LINEA() As String
            K = 0
            LINEA = sr.ReadLine.Split(",")
            Lat = (File.ReadAllLines(Archidir3).Count - 1)
            If Lat <> LINEA(5) Then
                MsgBox("El Numero de Registros No Conciden con la cantidad de registros (Linea de control)", MsgBoxStyle.Exclamation, "CONTAR COLUMNAS")
                Return False
            End If
            Do While (K < LINEA(5))
                K = K + 1
                Array_Consultas = sr.ReadLine.Split(",")
                If Array_Consultas.Length <> 25 Then
                    MsgBox("Error en la Linea " & K & " El Numero de campos No Conciden", MsgBoxStyle.Exclamation, "CONTAR COLUMNAS")
                    resul = False
                Else
                    resul = True
                End If
            Loop
            Return resul
        Catch ex As Exception
            Return Nothing
        End Try



    End Function
    Private Sub EmpezarValidacion(ByRef ruta As String)
        Dim InicioV As New Validacion1479
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

    Protected Sub ButtonValidar_Click(sender As Object, e As EventArgs) Handles ButtonValidar.Click
        Iniciar()
    End Sub
End Class