'Imports System.Data
Imports MySql.Data.MySqlClient
'Imports Excel = Microsoft.Office.Interop.Excel
'Imports System.Runtime.InteropServices
Public Class Pasar_ErroresRips
    Dim conect As New ClassConexion
    Dim conexion As String = conect.CrearConexion.ConnectionString
    Public Function Obtener_Errores_(ByVal op As String, ByVal id As String) As DataTable
        Try
            Dim tbl As New DataTable
            Dim coma As New MySqlDataAdapter
            Using cn As New MySqlConnection(conexion)
                cn.Open()
                Using cmd As New MySqlCommand("Pasarra_Errores", cn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Parameters.Add("op", MySqlDbType.Float).Value = op
                    cmd.Parameters.Add("id", MySqlDbType.VarChar).Value = id
                    cmd.CommandTimeout = 9000000
                    cmd.CommandType = CommandType.StoredProcedure
                    coma.SelectCommand = cmd
                    coma.Fill(tbl)
                    Return tbl
                    cn.Close()
                End Using
            End Using
        Catch ex As MySqlException
            Return Nothing
        End Try
    End Function
End Class
