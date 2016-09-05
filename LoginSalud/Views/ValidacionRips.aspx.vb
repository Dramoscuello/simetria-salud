Imports System.Data
Imports MySql.Data.MySqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports System.IO
Imports ClosedXML.Excel

Public Class ValidacionRips
    Inherits System.Web.UI.Page

    Dim claseprocedure As New Codprocedure
    Dim Nomre_Archivo As New DataTable

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Label3.Visible = False
    End Sub

    Private Sub cargar_Solo_Nombres(ByRef id As String, ByRef Archi2 As String, ByRef nombre As String)
        Dim controlCT As New DataTable
        controlCT = claseprocedure.ListarControl()
        Dim opcion As String
        For Each MiDataRow As DataRow In controlCT.Rows
            opcion = MiDataRow("Campo3")
            Select Case Mid(opcion, 1, 2).ToUpper
                Case "US"
                    Dim US_ As String() = Directory.GetFiles(Archi2, "US*")
                    If US_.Length = 0 Then
                        MsgBox("El archivo US no Existe-- " & opcion, MsgBoxStyle.Information, "Simetria Consolidated")
                        Label2.Text = ""
                        Exit Sub
                    End If
                    Label2.Text = "Importando Archivos de usuarios"
                    claseprocedure.RCargar_Control(Replace(US_(0), "\", "/"), "US", id)
                Case "AC"
                    Dim AC As String() = Directory.GetFiles(Archi2, "AC*")
                    If AC.Length = 0 Then
                        MsgBox("El archivo AC no Existe-- " & opcion, MsgBoxStyle.Information, "Simetria Consolidated")
                        Label2.Text = ""
                        Exit Sub
                    End If
                    Label2.Text = "Importando Archivos de Consultas"
                    claseprocedure.RCargar_Control(Replace(AC(0), "\", "/"), "AC", id)

                Case "AF"
                    Dim AF As String() = Directory.GetFiles(Archi2, "AF*")
                    If AF.Length = 0 Then
                        MsgBox("El archivo AF no Existe  - " & opcion, MsgBoxStyle.Information, "Simetria Consolidated")
                        Label2.Text = ""
                        Exit Sub
                    End If
                    Label2.Text = "Importando Archivos de transacciones"
                    claseprocedure.RCargar_Control(Replace(AF(0), "\", "/"), "AF", id)

                Case "AH"
                    Dim AH As String() = Directory.GetFiles(Archi2, "AH*")
                    If AH.Length = 0 Then
                        MsgBox("El archivo AH no Existe  - " & opcion, MsgBoxStyle.Information, "Simetria Consolidated")
                        Label2.Text = ""
                        Exit Sub
                    End If
                    Label2.Text = "Importando Archivos de Hospitalizacion"
                    claseprocedure.RCargar_Control(Replace(AH(0), "\", "/"), "AH", id)

                Case "AM"
                    Dim AM As String() = Directory.GetFiles(Archi2, "AM*")
                    If AM.Length = 0 Then
                        MsgBox("El archivo AM no Existe - " & opcion, MsgBoxStyle.Information, "Simetria Consolidated")
                        Label2.Text = ""
                        Exit Sub
                    End If
                    Label2.Text = "Importando Archivos de Medicamentos"
                    claseprocedure.RCargar_Control(Replace(AM(0), "\", "/"), "AM", id)
                Case "AN"
                    Dim AN As String() = Directory.GetFiles(Archi2, "AN*")
                    If AN.Length = 0 Then
                        MsgBox("El archivo AN no Existe - " & opcion, MsgBoxStyle.Information, "Simetria Consolidated")
                        Label2.Text = ""
                        Exit Sub
                    End If
                    Label2.Text = "Importando Archivos de Nacimiento"
                    claseprocedure.RCargar_Control(Replace(AN(0), "\", "/"), "AN", id)

                Case "AP"
                    Dim AP As String() = Directory.GetFiles(Archi2, "AP*")
                    If AP.Length = 0 Then
                        MsgBox("El archivo AP no Existe - " & opcion, MsgBoxStyle.Information, "Simetria Consolidated")
                        Label2.Text = ""
                        Exit Sub
                    End If
                    Label2.Text = "Importando Archivos de Procedimientos"
                    claseprocedure.RCargar_Control(Replace(AP(0), "\", "/"), "AP", id)

                Case "AT"
                    Dim AT As String() = Directory.GetFiles(Archi2, "AT*")
                    If AT.Length = 0 Then
                        MsgBox("El archivo AT no Existe - " & opcion, MsgBoxStyle.Information, "Simetria Consolidated")
                        Label2.Text = ""
                        Exit Sub
                    End If
                    Label2.Text = "Importando Archivos de Otros Servicios"
                    claseprocedure.RCargar_Control(Replace(AT(0), "\", "/"), "at01", id)
                Case "AU"
                    Dim AU As String() = Directory.GetFiles(Archi2, "AT*")
                    If AU.Length = 0 Then
                        MsgBox("El archivo AT no Existe - " & opcion, MsgBoxStyle.Information, "Simetria Consolidated")
                        Label2.Text = ""
                        Exit Sub
                    End If
                    Label2.Text = "Importando Archivos de Urgencias"
                    claseprocedure.RCargar_Control(Replace(AU(0), "\", "/"), "AU", id)

            End Select
        Next
        claseprocedure.Excluir()
        claseprocedure.Act_dATOSTB()
        claseprocedure.Act_edades_Q_E_V()
        claseprocedure.Act_CamposRep_()
        For Each MiDataRow As DataRow In controlCT.Rows
            opcion = MiDataRow("Campo3")
            Select Case Mid(opcion, 1, 2).ToUpper
                Case "AC"
                    claseprocedure.Validar_Consultas()
                    claseprocedure.Validar_Consultastari(DropDownListPorcentaje.Text, "CUPS")
                Case "AF"
                    claseprocedure.Validar_Transaccion()
                Case "AH"
                    claseprocedure.Validar_Hospitalizacion()
                Case "AM"
                    claseprocedure.Validar_Medicamentos()
                Case "AN"
                    claseprocedure.Validar_Nacimientos()
                Case "AP"
                    claseprocedure.Validar_Procedimientos()
                    claseprocedure.Validar_Procedimientostari(DropDownListPorcentaje.Text, "CUPS")
                Case "AT"
                    claseprocedure.Validar_Otros_servicios()
                    claseprocedure.Validar_Otros_serviciostari(DropDownListPorcentaje.Text, "CUPS")
                Case "AU"
                    claseprocedure.Validar_Urgencias()
                Case "US"
                    claseprocedure.Validar_Usuarios()
            End Select
        Next
        Label3.Visible = True
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
        Dim af_, am, ac, ah, an, ap, at, au, us, ok_ As New DataTable
        af_ = pasaerE.Obtener_Errores_AF("02") '.Rows.Count
        am = pasaerE.Obtener_Errores_AM("02")
        ac = pasaerE.Obtener_Errores_CA("02")
        ah = pasaerE.Obtener_Errores_AH("02")
        an = pasaerE.Obtener_Errores_AN("02")
        ap = pasaerE.Obtener_Errores_AP("02")
        at = pasaerE.Obtener_Errores_AT("02")
        au = pasaerE.Obtener_Errores_AU("02")
        us = pasaerE.Obtener_Errores_US("02")
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
            ClientScript.RegisterStartupScript(Me.GetType(), "alert", "alert('No hay Archivos para Descargar ');", True)

        End If
    End Sub

    Protected Sub ButtonInforme_Click(sender As Object, e As EventArgs) Handles ButtonInforme.Click
        Genera_Excel_errores()
    End Sub

    Protected Sub ButtonValidar_(sender As Object, e As EventArgs) Handles ButtonValidar.Click
        Dim conect As New ClassConexion
        Try
            If DropDownListPorcentaje.Text = "0" Then
                ClientScript.RegisterStartupScript(Me.GetType(), "alert", "alert('Debe seleccionar el porcentaje de validacion ');", True)
                Exit Sub
            End If
            claseprocedure.Eliminar_Registros_Usuarios("02")
            If IO.Directory.Exists(Server.MapPath("Validacion/")) Then
                For Each item As String In Directory.GetFiles(Server.MapPath("Validacion/"))
                    File.Delete(item)
                Next
            End If
            My.Computer.FileSystem.CreateDirectory(Server.MapPath("Validacion/"))

            Dim x As String = Server.MapPath("Validacion/")
            Dim path As String = FileUploadImportar.PostedFile.FileName
            Dim source As String = Replace(x, "\", "/")
            If Not String.IsNullOrEmpty(path) Then
                Dim ImageFiles As HttpFileCollection = Request.Files
                For i As Integer = 0 To ImageFiles.Count - 1
                    Dim file As HttpPostedFile = ImageFiles(i)
                    file.SaveAs(Server.MapPath("Validacion/") & file.FileName)
                    Label2.Text = "Cargando Archivos al Servidor"
                    Dim nomb As String = Mid(file.FileName, 1, 2)
                    If nomb = "CT" Or nomb = "Ct" Or nomb = "cT" Or nomb = "ct" Then
                        claseprocedure.RCargar_Control(source + file.FileName, UCase(nomb), "02")
                        Label2.Text = "Importando Archivos CT"
                    End If
                Next
                Dim ct As String() = Directory.GetFiles(source, "CT*")
                If ct.Length = 0 Then
                    ClientScript.RegisterStartupScript(Me.GetType(), "alert", "alert('Error el archivo CT no Existe  Verifique e intente nuevamente');", True)
                    Label2.Text = ""
                    Exit Sub
                Else
                    cargar_Solo_Nombres("02", source, "")
                End If
                My.Computer.FileSystem.DeleteFile(Server.MapPath("Validacion/") + path)

                path = Nothing
            ElseIf String.IsNullOrEmpty(path) Then
                ClientScript.RegisterStartupScript(Me.GetType(), "alert", "alert('Debe seleccionar los Archivos de Rips ');", True)
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


End Class