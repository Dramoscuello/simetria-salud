Imports System.IO
Imports MySql.Data.MySqlClient
Public Class Validacion1479
    Dim conect As New ClassConexion
    Dim conexion As String = conect.CrearConexion.ConnectionString


    Public Sub importarD1479(ByRef ArchiDir As String, ByRef IdUsuariA As String)
        Dim sSQL As String
        Dim Archidir3 As String
        Using cn As New MySqlConnection(conexion)
            cn.Open()
            Try
                Archidir3 = Replace(ArchiDir, "\", "/")
                sSQL = "SET AUTOCOMMIT = 0; LOAD DATA local INFILE '" & Archidir3 & "' INTO TABLE DBlevalidamos.importar1479 CHARACTER SET latin1 FIELDS TERMINATED BY ',' LINES TERMINATED BY '\n' IGNORE 1  LINES " &
                    "(C0,C1,C2,C3,C4,C5,C6,C7,C8,C9,C10,C11,C12,C13,C14,C15,C16,C17,C18,C19,C20,C21,C22,C23,C24,Usuario,FechaProceso) " &
                    "SET Usuario='" & IdUsuariA & "',FechaProceso='" & Format(CDate(Date.Now), "yyyyMMdd") & "'; COMMIT;"
                Dim oComando As New MySqlCommand(sSQL, cn)
                oComando.CommandType = CommandType.Text
                oComando.CommandTimeout = 5000000
                oComando.ExecuteNonQuery()
            Catch ex As MySqlException
                MsgBox(ex.Message)
            End Try
        End Using
    End Sub
    Public Sub Validar_Estructura(ByRef IdUsuariA As String)
        Try
            Using conn As New MySqlConnection(conexion)
                conn.Open()
                Dim oComando As New MySqlCommand("Validacion_Errores_Estructura_1479", conn)
                oComando.CommandType = CommandType.StoredProcedure
                oComando.Parameters.Add("USession", MySqlDbType.VarChar).Value = IdUsuariA
                oComando.CommandTimeout = 9999999
                oComando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub Validar_Detalles(ByRef IdUsuariA As String)
        Try
            Using conn As New MySqlConnection(conexion)
                conn.Open()
                Dim oComando As New MySqlCommand("Validacion_Errores_Detalle_1479", conn)
                oComando.CommandType = CommandType.StoredProcedure
                oComando.Parameters.Add("USession", MySqlDbType.VarChar).Value = IdUsuariA
                oComando.CommandTimeout = 9999999
                oComando.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub ActualizarEapb(ByRef IdUsuariA As String)
        With Me
            Dim sSQL As String = ""
            Dim res As Integer
            Try
                Dim cn As New MySqlConnection(conexion)
                cn.Open()
                sSQL = "Actualizar_Eapb"
                Using cmd As New MySqlCommand(sSQL, cn)
                    cmd.Parameters.Add("USession", MySqlDbType.VarChar).Value = IdUsuariA
                    cmd.Parameters.Add("OP", MySqlDbType.Float).Value = 0
                    cmd.CommandType = CommandType.StoredProcedure
                    res = cmd.ExecuteNonQuery()
                End Using
            Catch ex As MySqlException
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, sSQL)
            End Try
        End With
    End Sub
    Public Sub ImportarControl1479(ByRef ArchiDir As String, ByRef IdUsuariA As String)
        Dim Archidir3 As String
        Using conn As New MySqlConnection(conexion)
            Try
                Archidir3 = Replace(ArchiDir, "\", "/")
                Dim sr As New StreamReader(Archidir3)
                Dim Array_Consultas() As String
                Array_Consultas = sr.ReadLine.Split(",")
                'Linea = sr.ReadLine()
                Using cmd As New MySqlCommand("1479_Procedure", conn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Parameters.Add("op", MySqlDbType.VarChar).Value = "2"
                    cmd.Parameters.Add("NIT", MySqlDbType.VarChar).Value = "333"
                    cmd.Parameters.Add("C0", MySqlDbType.Int32).Value = Array_Consultas(0)
                    cmd.Parameters.Add("C1", MySqlDbType.VarChar).Value = Array_Consultas(1)
                    cmd.Parameters.Add("C2", MySqlDbType.Decimal).Value = Val(Array_Consultas(2))
                    cmd.Parameters.Add("C3", MySqlDbType.Date).Value = Array_Consultas(3)
                    cmd.Parameters.Add("C4", MySqlDbType.Date).Value = Array_Consultas(4)
                    cmd.Parameters.Add("C5", MySqlDbType.Decimal).Value = Array_Consultas(5)
                    cmd.Parameters.Add("USession", MySqlDbType.VarChar).Value = IdUsuariA
                    If conn.State = ConnectionState.Closed Then
                        conn.Open()
                    End If
                    cmd.ExecuteNonQuery()
                End Using
                sr.Close()
                conn.Close()
                'CType(Me.MdiParent, Mdi_Pricipal).Etiqueta.Text = "DATOS IMPORTADOS CORRECTAMENTE"
            Catch ex As Exception

            End Try
        End Using
    End Sub
    Public Sub Eliminar_registros(ByRef IdUsuariA As String)
        With Me
            Dim sSQL As String = "Eliminar_Registros1479"
            Try
                Using cn As New MySqlConnection(conexion)
                    cn.Open()
                    Using cmd As New MySqlCommand(sSQL, cn)
                        cmd.Parameters.Add("USession", MySqlDbType.VarChar).Value = IdUsuariA
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.ExecuteNonQuery()
                        cn.Close()
                    End Using
                End Using
            Catch ex As MySqlException

            End Try
        End With
    End Sub
    Public Function VerificarID(ByRef nit As String) As Integer
        Try
            Dim resultado As Integer
            Using cn As New MySqlConnection(conexion)
                cn.Open()
                Using cmd As New MySqlCommand("1479_Procedure", cn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandTimeout = 900000000
                    cmd.Parameters.Add("op", MySqlDbType.VarChar).Value = "1"
                    cmd.Parameters.Add("NIT", MySqlDbType.VarChar).Value = nit
                    cmd.Parameters.Add("C0", MySqlDbType.Float).Value = 0
                    cmd.Parameters.Add("C1", MySqlDbType.VarChar).Value = "1"
                    cmd.Parameters.Add("C2", MySqlDbType.Decimal).Value = 0
                    cmd.Parameters.Add("C3", MySqlDbType.Date).Value = CDate("2010-09-09")
                    cmd.Parameters.Add("C4", MySqlDbType.Date).Value = CDate("2010-09-09")
                    cmd.Parameters.Add("C5", MySqlDbType.Decimal).Value = 0
                    cmd.Parameters.Add("USession", MySqlDbType.VarChar).Value = "M"
                    resultado = Convert.ToInt32(cmd.ExecuteScalar())
                    If resultado = 0 Then
                        Return 0
                    Else
                        Return 1
                    End If
                End Using
                cn.Close()
            End Using
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
End Class
