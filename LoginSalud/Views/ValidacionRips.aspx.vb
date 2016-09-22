'Imports System.Data
'Imports MySql.Data.MySqlClient
'Imports Excel = Microsoft.Office.Interop.Excel
'Imports System.Runtime.InteropServices
Imports System.IO
Imports ClosedXML.Excel
Imports System.Threading

Imports LoginSalud.Controllers
Public Class ValidacionRips
    Inherits System.Web.UI.Page
    Dim idusu As String

    Dim Nomre_Archivo As New DataTable
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

    Private Sub cargar_Solo_Nombres(ByRef id As String, ByRef Archi2 As String, ByRef nombre As String)
        Dim controlCT As New DataTable
        Dim claseprocedure As New RipsController

        controlCT = claseprocedure.ListarControl(id)
        Dim opcion As String
        For Each MiDataRow As DataRow In controlCT.Rows
            opcion = MiDataRow("Campo3")
            Select Case Mid(opcion, 1, 2).ToUpper
                Case "US"
                    Dim US_ As String() = Directory.GetFiles(Archi2, "US*")
                    If US_.Length = 0 Then
                        ClientScript.RegisterStartupScript(Me.GetType, "alerta", "swal('¡Alerta!','¡El archivo  " & opcion & " no Existe!', 'warning')", True)
                        Exit Sub
                    End If
                    claseprocedure.RCargar_Control(Replace(US_(0), "\", "/"), "US", id)
                Case "AC"
                    Dim AC As String() = Directory.GetFiles(Archi2, "AC*")
                    If AC.Length = 0 Then
                        ClientScript.RegisterStartupScript(Me.GetType, "alerta", "swal('¡Alerta!','¡El archivo " & opcion & " no Existe!', 'warning')", True)
                        Exit Sub
                    End If
                    claseprocedure.RCargar_Control(Replace(AC(0), "\", "/"), "AC", id)
                Case "AF"
                    Dim AF As String() = Directory.GetFiles(Archi2, "AF*")
                    If AF.Length = 0 Then
                        ClientScript.RegisterStartupScript(Me.GetType, "alerta", "swal('¡Alerta!','¡El archivo  " & opcion & " no Existe!', 'warning')", True)
                        Exit Sub
                    End If
                    claseprocedure.RCargar_Control(Replace(AF(0), "\", "/"), "AF", id)
                Case "AH"
                    Dim AH As String() = Directory.GetFiles(Archi2, "AH*")
                    If AH.Length = 0 Then
                        ClientScript.RegisterStartupScript(Me.GetType, "alerta", "swal('¡Alerta!','¡El archivo  " & opcion & " no Existe!', 'warning')", True)
                        Exit Sub
                    End If
                    claseprocedure.RCargar_Control(Replace(AH(0), "\", "/"), "AH", id)
                Case "AM"
                    Dim AM As String() = Directory.GetFiles(Archi2, "AM*")
                    If AM.Length = 0 Then
                        ClientScript.RegisterStartupScript(Me.GetType, "alerta", "swal('¡Alerta!','¡El archivo " & opcion & " no Existe!', 'warning')", True)
                        Exit Sub
                    End If
                    claseprocedure.RCargar_Control(Replace(AM(0), "\", "/"), "AM", id)
                Case "AN"
                    Dim AN As String() = Directory.GetFiles(Archi2, "AN*")
                    If AN.Length = 0 Then
                        ClientScript.RegisterStartupScript(Me.GetType, "alerta", "swal('¡Alerta!','¡El archivo  " & opcion & " no Existe!', 'warning')", True)
                        Exit Sub
                    End If
                    claseprocedure.RCargar_Control(Replace(AN(0), "\", "/"), "AN", id)
                Case "AP"
                    Dim AP As String() = Directory.GetFiles(Archi2, "AP*")
                    If AP.Length = 0 Then
                        ClientScript.RegisterStartupScript(Me.GetType, "alerta", "swal('¡Alerta!','¡El archivo  " & opcion & " no Existe!', 'warning')", True)
                        Exit Sub
                    End If
                    claseprocedure.RCargar_Control(Replace(AP(0), "\", "/"), "AP", id)
                Case "AT"
                    Dim AT As String() = Directory.GetFiles(Archi2, "AT*")
                    If AT.Length = 0 Then
                        ClientScript.RegisterStartupScript(Me.GetType, "alerta", "swal('¡Alerta!','¡El archivo " & opcion & " no Existe!', 'warning')", True)
                        Exit Sub
                    End If
                    claseprocedure.RCargar_Control(Replace(AT(0), "\", "/"), "at01", id)
                Case "AU"
                    Dim AU As String() = Directory.GetFiles(Archi2, "AU*")
                    If AU.Length = 0 Then
                        ClientScript.RegisterStartupScript(Me.GetType, "alerta", "swal('¡Alerta!','¡El archivo " & opcion & " no Existe!', 'warning')", True)
                        Exit Sub
                    End If
                    claseprocedure.RCargar_Control(Replace(AU(0), "\", "/"), "AU", id)
            End Select
        Next

        Dim eXC As String = claseprocedure.Excluir("11", idusu).ToString
        claseprocedure.Act_dATOSTB(idusu)
        claseprocedure.Act_edades_Q_E_V(idusu)
        claseprocedure.Act_CamposRep(idusu, eXC)
        For Each MiDataRow As DataRow In controlCT.Rows
            opcion = MiDataRow("Campo3")
            Select Case Mid(opcion, 1, 2).ToUpper
                Case "AC"

                    claseprocedure.Validar_Consultas(idusu, eXC, DropDownListPorcentaje.Text)
                   ' claseprocedure.Validar_Consultastari(idusu, DropDownListPorcentaje.Text, "CUPS")
                Case "AF"
                    claseprocedure.Validar_Transaccion(idusu, eXC)
                Case "AH"
                    claseprocedure.Validar_Hospitalizacion(idusu, eXC)
                Case "AM"
                    claseprocedure.Validar_Medicamentos(idusu, eXC)
                Case "AN"
                    claseprocedure.Validar_Nacimientos(idusu, eXC)
                Case "AP"
                    claseprocedure.Validar_Procedimientos(idusu, eXC, DropDownListPorcentaje.Text)
                  ''  claseprocedure.Validar_Procedimientostari(idusu, DropDownListPorcentaje.Text, "CUPS")
                Case "AT"
                    claseprocedure.Validar_Otros_servicios(idusu, eXC, DropDownListPorcentaje.Text)
                  '  claseprocedure.Validar_Otros_serviciostari(idusu, DropDownListPorcentaje.Text, "CUPS")
                Case "AU"
                    claseprocedure.Validar_Urgencias(idusu, eXC)
                Case "US"
                    claseprocedure.Validar_Usuarios(idusu, eXC)
            End Select
        Next
        ButtonInforme.Enabled = True
        'Thread.Sleep(2000)
        ClientScript.RegisterStartupScript(Me.GetType, "ok", "swal('¡Validación terminada!','¡descargue el informe!', 'success')", True)
    End Sub
    Private Sub Llenar_Grid()
        'Dim dt As New DataTable()
        'dt.Columns.Add("Nombre Archivo")
        'dt.Columns.Add("Numero de Registros")
        'dt.Columns.Add("Registros Erroneos")
        ''  GridViewDatos.DataSource =
        'Dim llenar_grid_temp As New DataSet
        'llenar_grid_temp = claseprocedure.Llenar
        'Dim row As DataRow = dt.NewRow()
        'Dim fi As String = claseprocedure.nombre_fichero()
        'Dim fich As String = Replace(fi, ".txt", ".xlsx")
        'row("Nombre Archivo") = fich
        'row("Numero de Registros") = llenar_grid_temp.Tables(0).Rows(0).Item(0).ToString()
        'row("Registros Erroneos") = llenar_grid_temp.Tables(0).Rows(0).Item(1).ToString()
        'dt.Rows.Add(row)
        'GridViewResultado.DataSource = dt
        'GridViewResultado.DataBind()
    End Sub


    Dim pasaerE As New Pasar_ErroresRips
    Sub Genera_Excel_errores()
        'If String.IsNullOrEmpty(FileUploadImportar.PostedFile.FileName) Then
        '    ClientScript.RegisterStartupScript(Me.GetType(), "alert", "alert('Debe Primero Validar los Archivos de Rips ');", True)
        '    Exit Sub
        'End If
        Try
            Dim af_, am, ac, ah, an, ap, at, au, us, ok_ As New DataTable
            af_ = pasaerE.Obtener_Errores_("1", idusu)
            ac = pasaerE.Obtener_Errores_("2", idusu)
            ah = pasaerE.Obtener_Errores_("3", idusu)
            am = pasaerE.Obtener_Errores_("4", idusu)
            an = pasaerE.Obtener_Errores_("5", idusu)
            ap = pasaerE.Obtener_Errores_("6", idusu)
            at = pasaerE.Obtener_Errores_("7", idusu)
            au = pasaerE.Obtener_Errores_("8", idusu)
            us = pasaerE.Obtener_Errores_("9", idusu)
            'totalfac = pasaerE.Obtener_Errores_("10", idusu)
            Dim wb As New XLWorkbook()
            If af_.Rows.Count > 0 Then
                wb.Worksheets.Add(af_, "Errores_AF")
            End If
            If am.Rows.Count > 0 Then
                wb.Worksheets.Add(am, "Errores_Medicamentos")
            End If

            If ac.Rows.Count > 0 Then
                wb.Worksheets.Add(ac, "Errores_en_Consultas")
            End If

            If ah.Rows.Count > 0 Then
                wb.Worksheets.Add(ah, "Errores_en_Hospitali")
            End If

            If an.Rows.Count > 0 Then
                wb.Worksheets.Add(an, "Errores_en_Nacimiento")
            End If
            If ap.Rows.Count > 0 Then
                wb.Worksheets.Add(ap, "Errores_en_Procedimientos")
            End If

            If at.Rows.Count > 0 Then
                wb.Worksheets.Add(at, "Errores_en_OtrosServicios")
            End If

            If au.Rows.Count > 0 Then
                wb.Worksheets.Add(au, "Errores_en_Urgencias")
            End If

            If us.Rows.Count > 0 Then
                wb.Worksheets.Add(us, "Errores_en_Usuarios")
            End If
            ok_.Columns.Add(" ")
            ok_.Rows.Add(" ")
            wb.Worksheets.Add(ok_, ".")
            wb.Author = "Simetria"
            Response.Clear()
            Response.Buffer = True
            Response.Charset = ""
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            Response.AddHeader("content-disposition", "attachment;filename=Errores_Rips.xlsx")
            Dim MyMemoryStream As New MemoryStream()
            wb.SaveAs(MyMemoryStream)
            MyMemoryStream.WriteTo(Response.OutputStream)
            Response.Flush()
            Response.End()
            If af_.Rows.Count = 0 And am.Rows.Count And ac.Rows.Count = 0 And ah.Rows.Count = 0 And an.Rows.Count = 0 And ap.Rows.Count = 0 And at.Rows.Count = 0 And au.Rows.Count = 0 And us.Rows.Count = 0 Then
            Else
                ClientScript.RegisterStartupScript(Me.GetType, "error", "swal('¡Error!','No hay Archivos para Descargar', 'error')", True)
            End If
        Catch ex As Exception
            If ex.InnerException Is Nothing Then
                ClientScript.RegisterStartupScript(Me.GetType, "error1", "swal('¡Error!'," + ex.Message.ToString() + ", 'error')", True)
            Else
                ClientScript.RegisterStartupScript(Me.GetType, "error2", "swal('¡Error!'," + ex.InnerException.Message.ToString() + ", 'error')", True)
            End If
        End Try
    End Sub

    Protected Sub ButtonInforme_Click(sender As Object, e As EventArgs) Handles ButtonInforme.Click
        Genera_Excel_errores()
    End Sub
    Dim Error_V2 As String
    Public Function errorValidacion(ByRef Error_V As String)
        ClientScript.RegisterStartupScript(Me.GetType, "alerta", "swal('Error " & Error_V & "', 'warning')", True)
        Error_V2 = Error_V
        Return Error_V2
    End Function



    Protected Sub ButtonValidar_(sender As Object, e As EventArgs) Handles ButtonValidar.Click
        Dim conect As New ClassConexion
        Dim claseprocedure As New CodRips

        If Error_V2 IsNot Nothing Then
            Exit Sub
        End If
        Try
            claseprocedure.Eliminar_Registros_Usuarios(idusu)
            If DropDownListPorcentaje.Text = "0" Then
                ClientScript.RegisterStartupScript(Me.GetType, "alerta", "swal('¡Alerta!','Debe seleccionar el porcentaje de validación', 'warning')", True)
                Exit Sub
            End If

            If IO.Directory.Exists(Server.MapPath(idusu & "/")) Then
                For Each item As String In Directory.GetFiles(Server.MapPath(idusu & "/"))
                    File.Delete(item)
                Next
            End If
            My.Computer.FileSystem.CreateDirectory(Server.MapPath(idusu & "/"))

            Dim x As String = Server.MapPath(idusu & "/")
            Dim path As String = FileUploadImportar.PostedFile.FileName
            Dim source As String = Replace(x, "\", "/")
            If Not String.IsNullOrEmpty(path) Then

                Dim ImageFiles As HttpFileCollection = Request.Files
                For i As Integer = 0 To ImageFiles.Count - 1
                    Dim file As HttpPostedFile = ImageFiles(i)
                    file.SaveAs(Server.MapPath(idusu & "/") & file.FileName)
                    Dim nomb As String = Mid(file.FileName, 1, 2)
                    If nomb = "CT" Or nomb = "Ct" Or nomb = "cT" Or nomb = "ct" Then
                        claseprocedure.RCargar_Control(source + file.FileName, UCase(nomb), idusu)
                    End If
                Next
                Dim ct As String() = Directory.GetFiles(source, "CT*")
                If ct.Length = 0 Then
                    ClientScript.RegisterStartupScript(Me.GetType, "error3", "swal('¡Error!','Error el archivo CT no Existe  Verifique e intente nuevamente', 'error')", True)
                    Exit Sub
                Else
                    cargar_Solo_Nombres(idusu, source, "")
                End If
                My.Computer.FileSystem.DeleteFile(Server.MapPath(idusu & "/") + path)
                path = Nothing
            ElseIf String.IsNullOrEmpty(path) Then
                ClientScript.RegisterStartupScript(Me.GetType, "alerta", "swal('¡Alerta!','Debe seleccionar los Archivos de Rips', 'warning')", True)
            End If
        Catch ex As Exception
            ClientScript.RegisterStartupScript(Me.GetType, "alerta", "swal('¡Alerta!',' Error " & ex.Message.ToString & " ', 'warning')", True)

            Exit Sub
            'If ex.InnerException Is Nothing Then

            '    ClientScript.RegisterStartupScript(Me.GetType, "error4", "swal('¡Error!'," + ex.Message.ToString() + ", 'error')", True)
            'Else
            '    ClientScript.RegisterStartupScript(Me.GetType, "error5", "swal('¡Error!'," + ex.InnerException.Message.ToString() + ", 'error')", True)
            'End If
        End Try
    End Sub
End Class