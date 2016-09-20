'Imports System.Data
Imports MySql.Data.MySqlClient
'Imports Excel = Microsoft.Office.Interop.Excel
'Imports System.Runtime.InteropServices


Public Class CodRips
    Dim conect As New ClassConexion
    Dim oComando As MySqlCommand

    Dim conexion As String = conect.CrearConexion.ConnectionString
    Public Sub Eliminar_Registros_Usuarios(ByRef id As String)
        With Me
            Dim sSQL As String = ""
            Try
                Dim Conectar_ As New MySqlConnection(conexion)
                Conectar_.Open()
                sSQL = "Eliminar_Registros_Usuarios"
                Using cmd As New MySqlCommand(sSQL, Conectar_)
                    cmd.Parameters.Add("USession", MySqlDbType.VarChar).Value = id
                    cmd.CommandTimeout = 900000000
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.ExecuteNonQuery()
                End Using
            Catch ex As MySqlException
                MsgBox(ex.Message, , sSQL)
            End Try
        End With
    End Sub


    Public Function Llenar() As DataSet
        Try
            Dim myData As New DataSet
            Dim myAdapter As New MySqlDataAdapter
            Dim Conectar_ As New MySqlConnection(conexion)
            Conectar_.Open()
            Dim cmd As New MySqlCommand
            cmd.Connection = Conectar_
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "Estado_Archivo"
            myAdapter.SelectCommand = cmd
            myAdapter.Fill(myData)
            Return myData
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        End Try

    End Function

    Public Function ListarControl(ByRef id_u As String) As DataTable
        Try
            Dim myData As New DataTable
            Dim myAdapter As New MySqlDataAdapter
            Using cn As New MySqlConnection(conexion)
                cn.Open()
                ssql = "SELECT Campo3 FROM CT where Usuario='" & id_u & "'"
                oComando = New MySqlCommand(ssql, cn)
                oComando.CommandType = CommandType.Text
                oComando.CommandTimeout = 5000000
                myAdapter.SelectCommand = oComando
                myAdapter.Fill(myData)
                cn.Close()
            End Using
            Return myData
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        End Try
    End Function
    Dim ssql As String
    Public Sub TruncateControl()
        Try
            Using cn As New MySqlConnection(conexion)
                cn.Open()
                ssql = "TRUNCATE ct"
                oComando = New MySqlCommand(ssql, cn)
                oComando.CommandType = CommandType.Text
                oComando.CommandTimeout = 5000000
                oComando.ExecuteNonQuery()
                cn.Close()
            End Using
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub RCargar_Control(ByRef Archi2 As String, ByRef ntba As String, ByRef IdUsuariA As String)
        Try

            Using cn As New MySqlConnection(conexion)
                If cn.State = ConnectionState.Closed Then
                    cn.Open()
                End If

                Select Case ntba

                    Case "US"
                        ssql = "SET AUTOCOMMIT=0; LOAD DATA LOCAL INFILE '" & Archi2 & "' INTO TABLE DBlevalidamos." & ntba & " CHARACTER SET latin1 FIELDS TERMINATED BY ',' LINES TERMINATED BY '\r\n' " &
                              "(Campo1,Campo2,Campo3,Campo4,Campo5,Campo6,Campo7,Campo8,Campo9,Campo10,Campo11,Campo12,Campo13,Campo14,Campo15,Prestador,Atenciones,Regimen,Entidad,@fecha_afil,DX,DESCRIPCION,CUOTAMODERADORA,Usuario,@FechaNacimiento)" &
                              "SET fecha_afil=str_to_date(@fecha_afil,'%d/%m/%Y'), FechaNacimiento=str_to_date(@FechaNacimiento,'%d/%m/%Y'), Usuario='" & IdUsuariA & "'; COMMIT;"

                    Case "CT"
                        ssql = "SET AUTOCOMMIT=0; LOAD DATA LOCAL INFILE '" & Archi2 & "' INTO TABLE DBlevalidamos." & ntba & " FIELDS TERMINATED BY ',' LINES TERMINATED BY '\r\n' " &
                              "(Campo1,@Campo2, Campo3, Campo4,Campo5,Num_Radicacion, id,IdRecepcio,Campo0, RS,@Usuario) " &
                              "SET Campo2=str_to_date(@Campo2,'%d/%m/%Y'), Usuario='" & IdUsuariA & "'; COMMIT;"
                    Case "AF"
                        ssql = "SET AUTOCOMMIT=0; LOAD DATA LOCAL INFILE '" & Archi2 & "' INTO TABLE DBlevalidamos." & ntba & " CHARACTER SET latin1 FIELDS TERMINATED BY ',' LINES TERMINATED BY '\n' " &
                               "(Campo1,Campo2 ,Campo3 ,Campo4 , Campo5 , @Campo6 ,@Campo7 ,@Campo8 ,Campo9 , Campo10 ,Campo11 ,Campo12 ,Campo13 ,Campo14 , Campo15, Campo16 , Campo17	, n_id , Num_Radicacion , id , IdRecepcio ,DX , RS, @Usuario ) " &
                               "SET Campo6=str_to_date(@Campo6,'%d/%m/%Y'), Campo7=str_to_date(@Campo7,'%d/%m/%Y'), Campo8=STR_To_DATE(@Campo8,'%d/%m/%Y'), Usuario='" & IdUsuariA & "'; COMMIT;"

                    Case "AC"
                        ssql = "SET AUTOCOMMIT=0; LOAD DATA LOCAL INFILE '" & Archi2 & "' INTO TABLE DBlevalidamos." & ntba & " CHARACTER SET latin1 FIELDS TERMINATED BY ',' LINES TERMINATED BY '\r\n' " &
                            "(Campo1,Campo2,Campo3,Campo4,@Campo5,Campo6,Campo7,Campo8,Campo9,Campo10,Campo11,Campo12,Campo13,Campo14,Campo15,Campo16,Campo17,Campo18,EAPB,Tipo_Usuario,Edad,U_Edad,Sexo,Cod_Dpto,Cod_Mun,Cod_Zona,Num_Contrato,Plandebeneficios,Num_Poliza,EdadEtareo,EdadVigilancia,EdadQuinquenio,Entidad,Regimen,@fecha_afil,@FechaNacimiento,Usuario) " &
                            "SET Campo5=str_to_date(@Campo5,'%d/%m/%Y'), fecha_afil=str_to_date(@fecha_afil,'%d/%m/%Y'), FechaNacimiento=str_to_date(@FechaNacimiento,'%d/%m/%Y'), Usuario='" & IdUsuariA & "'; COMMIT;"

                    Case "AH"
                        ssql = "SET AUTOCOMMIT=0; LOAD DATA LOCAL INFILE '" & Archi2 & "' INTO TABLE DBlevalidamos." & ntba & " CHARACTER SET latin1 FIELDS TERMINATED BY ',' LINES TERMINATED BY '\r\n' " &
                           "(Campo1,Campo2,Campo3,Campo4,Campo5,@Campo6,Campo7,Campo8,Campo9,Campo10,Campo11,Campo12,Campo13,Campo14,Campo15,Campo16,Campo17,@Campo18,Campo19,Campo20,EAPB,Tipo_Usuario,Edad,U_Edad,Sexo,Cod_Dpto,Cod_Mun,Cod_Zona,Num_Contrato,Plandebeneficios,Num_Poliza,EdadEtareo,EdadVigilancia,EdadQuinquenio,Entidad,Regimen,@fecha_afil,@FechaNacimiento,Usuario) " &
                           "SET Campo6=str_to_date(@Campo6,'%d/%m/%Y'), fecha_afil=str_to_date(@fecha_afil,'%d/%m/%Y'), Campo18=str_to_date(@Campo18,'%d/%m/%Y'), FechaNacimiento=str_to_date(@FechaNacimiento,'%d/%m/%Y'), Usuario='" & IdUsuariA & "'; COMMIT;"

                    Case "AM"
                        ssql = "SET AUTOCOMMIT=0; LOAD DATA LOCAL INFILE '" & Archi2 & "' INTO TABLE DBlevalidamos." & ntba & " CHARACTER SET latin1 FIELDS TERMINATED BY ',' LINES TERMINATED BY '\r\n' " &
                                 "(Campo1,Campo2,Campo3,Campo4,Campo5,Campo6,Campo7,Campo8,Campo9,Campo10,Campo11,Campo12,Campo13,Campo14,Campo15,FECHA,EAPB,Tipo_Usuario,Edad,U_Edad,Sexo,Cod_Dpto,Cod_Mun,Cod_Zona,Num_Contrato,Plandebeneficios,Num_Poliza,EdadEtareo,EdadVigilancia,EdadQuinquenio,Entidad,Regimen,@fecha_afil,FechaNacimiento,Usuario) " &
                                 "SET fecha_afil=str_to_date(@fecha_afil,'%d/%m/%Y'), FechaNacimiento=str_to_date(@FechaNacimiento,'%d/%m/%Y'), Usuario='" & IdUsuariA & "'; COMMIT;"

                    Case "AN"
                        ssql = "SET AUTOCOMMIT=0; LOAD DATA LOCAL INFILE '" & Archi2 & "' INTO TABLE DBlevalidamos." & ntba & " CHARACTER SET latin1 FIELDS TERMINATED BY ',' LINES TERMINATED BY '\r\n' " &
                           "(Campo1,Campo2,Campo3,Campo4,@Campo5,Campo6,Campo7,Campo8,Campo9,Campo10,Campo11,Campo12,@Campo13,Campo14,Campo15,EAPB,Tipo_Usuario,Edad,U_Edad,Sexo,Cod_Dpto,Cod_Mun,Cod_Zona,Num_Contrato,Plandebeneficios,Num_Poliza,EdadEtareo,EdadVigilancia,EdadQuinquenio,Entidad,Regimen,@fecha_afil,@FechaNacimiento,Usuario) " &
                           "SET Campo5=str_to_date(@Campo5,'%d/%m/%Y'), Campo13=str_to_date(@Campo13,'%d/%m/%Y'),fecha_afil=str_to_date(@fecha_afil,'%d/%m/%Y'), FechaNacimiento=str_to_date(@FechaNacimiento,'%d/%m/%Y'), Usuario='" & IdUsuariA & "'; COMMIT;"

                    Case "AP"
                        ssql = "SET AUTOCOMMIT=0; LOAD DATA LOCAL INFILE '" & Archi2 & "' INTO TABLE DBlevalidamos." & ntba & " CHARACTER SET latin1 FIELDS TERMINATED BY ',' LINES TERMINATED BY '\r\n' " &
                               "(Campo1,Campo2,Campo3,Campo4,@Campo5,Campo6,Campo7,Campo8,Campo9,Campo10,Campo11,Campo12,Campo13,Campo14,Campo15,Campo16,EAPB,Tipo_Usuario,Edad,U_Edad,Sexo,Cod_Dpto,Cod_Mun,Cod_Zona,Num_Contrato,Plandebeneficios,Num_Poliza,EdadEtareo,EdadVigilancia,EdadQuinquenio,Entidad,Regimen,@fecha_afil,@FechaNacimiento,Usuario) " &
                               "SET Campo5=str_to_date(@Campo5,'%d/%m/%Y'), fecha_afil=str_to_date(@fecha_afil,'%d/%m/%Y'), FechaNacimiento=str_to_date(@FechaNacimiento,'%d/%m/%Y'), Usuario='" & IdUsuariA & "' ; COMMIT;"

                    Case "at01"
                        ssql = "SET AUTOCOMMIT=0; LOAD DATA LOCAL INFILE '" & Archi2 & "' INTO TABLE DBlevalidamos." & ntba & " CHARACTER SET latin1 FIELDS TERMINATED BY ',' LINES TERMINATED BY '\r\n' " &
                               "(Campo1,Campo2,Campo3,Campo4,Campo5,Campo6,Campo7,Campo8,Campo9,Campo10,Campo11,Campo12,Campo13,FECHA,EAPB,Tipo_Usuario,Edad,U_Edad,Sexo,Cod_Dpto,Cod_Mun,Cod_Zona,Num_Contrato,Plandebeneficios,Num_Poliza,EdadEtareo,EdadVigilancia,EdadQuinquenio,Entidad,Regimen,@fecha_afil,@FechaNacimiento,Usuario) " &
                               "SET fecha_afil=str_to_date(@fecha_afil,'%d/%m/%Y'), FechaNacimiento=str_to_date(@FechaNacimiento,'%d/%m/%Y'), Usuario='" & IdUsuariA & "'; COMMIT;"

                    Case "AU"
                        ssql = "SET AUTOCOMMIT=0; LOAD DATA LOCAL INFILE '" & Archi2 & "' INTO TABLE DBlevalidamos." & ntba & " CHARACTER SET latin1 FIELDS TERMINATED BY ',' LINES TERMINATED BY '\r\n' " &
                          "(Campo1,Campo2,Campo3,Campo4,@Campo5,Campo6,Campo7,Campo8,Campo9,Campo10,Campo11,Campo12,Campo13,Campo14,Campo15,@Campo16,Campo17,EAPB,Tipo_Usuario,Edad,U_Edad,Sexo,Cod_Dpto,Cod_Mun,Cod_Zona,Num_Contrato,Plandebeneficios,Num_Poliza,EdadEtareo,EdadVigilancia,EdadQuinquenio,Entidad,Regimen,@fecha_afil,@FechaNacimiento,Usuario) " &
                          "SET Campo5=str_to_date(@Campo5,'%d/%m/%Y'), fecha_afil=str_to_date(@fecha_afil,'%d/%m/%Y'), Campo16=str_to_date(@Campo16,'%d/%m/%Y'), FechaNacimiento=str_to_date(@FechaNacimiento,'%d/%m/%Y'), Usuario='" & IdUsuariA & "'; COMMIT;"
                End Select
                oComando = New MySqlCommand(ssql, cn)
                oComando.CommandType = CommandType.Text
                oComando.CommandTimeout = 5000000
                oComando.ExecuteNonQuery()

            End Using

        Catch ex As Exception

        Finally
        End Try
    End Sub
    Dim iduser As New ValidacionRips

    Public Sub Act_dATOSTB(ByRef id As String)
        Using cn As New MySqlConnection(conexion)
            cn.Open()

            Try
                Using cmd As New MySqlCommand("Act_Datos_", cn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandTimeout = 900000000
                    cmd.Parameters.Add("USession", MySqlDbType.VarChar).Value = id
                    cmd.ExecuteNonQuery()
                End Using
            Catch ex As MySqlException
                MsgBox(ex.Message, MsgBoxStyle.Information,)
            End Try
        End Using
    End Sub
    Public Sub Act_edades_Q_E_V(ByRef id As String)
        Dim sSQL As String
        Using cn As New MySqlConnection(conexion)
            cn.Open()
            sSQL = "Edad_Q_E_V_1"
            Try
                Using cmd As New MySqlCommand(sSQL, cn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Parameters.Add("USession", MySqlDbType.VarChar).Value = id
                    cmd.CommandTimeout = 900000000
                    cmd.ExecuteNonQuery()
                End Using
            Catch ex As MySqlException
                MsgBox(ex.Message, , sSQL)
            End Try
        End Using
    End Sub
    Public Sub Act_CamposRep(ByRef id As String, ByRef Excluir As String)
        Dim sSQL As String = ""
        Try
            sSQL = "Act_CamposRep"
            Using cn As New MySqlConnection(conexion)
                cn.Open()
                Using cmd As New MySqlCommand(sSQL, cn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Parameters.Add("USession", MySqlDbType.VarChar).Value = id
                    cmd.Parameters.Add("Excluir", MySqlDbType.Float).Value = Excluir
                    cmd.CommandTimeout = 9000000
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As MySqlException

        End Try

    End Sub
    Public Function Excluir(ByRef op As String, ByRef ID As String) As String
        Try
            Dim PExcluir As String = ""
            Dim tbl As New DataTable

            Try
                Using Conectar_ As New MySqlConnection(conexion)


                    If Conectar_.State = ConnectionState.Closed Then
                        Conectar_.Open()
                    End If
                    Using cmd As New MySqlCommand("Pasarra_Errores", Conectar_)
                        cmd.CommandTimeout = 900000000
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.Parameters.Add("op", MySqlDbType.VarChar).Value = op
                        cmd.Parameters.Add("id", MySqlDbType.VarChar).Value = ID
                        Dim dt As New DataTable()
                        Dim da As New MySqlDataAdapter(cmd)
                        If da.Fill(tbl) > 0 Then
                            If tbl.Rows(0)("CUS") > 1 And tbl.Rows(0)("CAF") = 1 Then
                                PExcluir = 1
                            Else
                                PExcluir = 0
                            End If
                        End If
                        Conectar_.Close()
                    End Using
                End Using
                Return PExcluir
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Information, "Simetria Consolidated")
                Return Nothing
            End Try
        Catch ex As MySqlException
            Return Nothing
        End Try
    End Function




    Public Sub Validar_Consultas(ByRef id As String, ByRef PExcluir As String, ByRef porce As String)
        Dim sSQL As String

        sSQL = "ERRORES_EN_CONSULTA"
        Try
            Using cn As New MySqlConnection(conexion)
                cn.Open()

                Using cmd As New MySqlCommand(sSQL, cn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandTimeout = 900000000
                    cmd.Parameters.Add("USession", MySqlDbType.VarChar).Value = id
                    cmd.Parameters.Add("Excluir", MySqlDbType.Float).Value = PExcluir
                    cmd.Parameters.Add("porce", MySqlDbType.Float).Value = CInt(porce)
                    cmd.ExecuteNonQuery()
                    cn.Close()
                End Using
            End Using
        Catch ex As MySqlException
            MsgBox(ex.Message, , sSQL)
        End Try
    End Sub
    Dim porce, por1 As Integer
    Dim tari As String
    'Public Sub Validar_Consultastari(ByRef id As String, ByRef Porcentajetxt As String, ByRef CbTipoTarifa As String)
    '    Dim cs As String, SSQL As String

    '    por1 = CInt(Porcentajetxt)
    '    cs = ""
    '    SSQL = ""
    '    porce = CInt(Porcentajetxt)
    '    tari = CbTipoTarifa
    '    Dim iResultado As Integer
    '    Try
    '        Using conn As New MySqlConnection(conexion)
    '            conn.Open()

    '            SSQL = "SET AUTOCOMMIT = 0; INSERT INTO error_ac_ (TIPO_IDENTIFI, NUM_IDENTIFI, NUM_FACTURA, FECHA_CONSULTA, CODIGO_CONS, DX_PPAL, DESCRIPCION_DEL_ERROR, ERROR1,USUARIO) SELECT ac.Campo3,ac.Campo4,ac.Campo1, ac.Campo5, ac.Campo7, ac.Campo10,'DIFERENCIA DE TARIFA' AS DESCRIP,ac.Campo15-ROUND(t.VALOR *(" & porce & " / 100), -2) AS PORC,'" & id & "' FROM ac INNER JOIN tarifas_1 t ON t.CÓDIGO=ac.Campo7 AND t.AÑO = YEAR(ac.Campo5) WHERE t.MANUAL = 'cups' AND ROUND(t.VALOR *(" & porce & " / 100), 1)<>ac.Campo10 AND (ac.Campo15-ROUND(t.VALOR *(" & porce & " / 100), 1))>0 AND ac.Usuario='" & id & "'; COMMIT;"
    '            Dim cmd As New MySqlCommand(SSQL, conn)
    '            cmd.CommandType = CommandType.Text
    '            cmd.CommandTimeout = 9000000
    '            iResultado = cmd.ExecuteNonQuery()
    '            ' ejecutar comando 
    '        End Using
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Exclamation, "VALIDANDO TARIFAS EN CONSULTAS ")
    '    End Try
    'End Sub
    Public Sub Validar_Hospitalizacion(ByRef id As String, ByRef PExcluir As String)
        Dim sSQL As String
        sSQL = "ERRORES_EN_HOSPITALIZACION"
        Try
            Using cn As New MySqlConnection(conexion)
                cn.Open()
                Using cmd As New MySqlCommand(sSQL, cn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandTimeout = 900000000
                    cmd.Parameters.Add("USession", MySqlDbType.VarChar).Value = id
                    cmd.Parameters.Add("Excluir", MySqlDbType.Float).Value = PExcluir
                    cmd.ExecuteNonQuery()
                    cn.Close()
                End Using
            End Using
        Catch ex As MySqlException
            MsgBox(ex.Message, , sSQL)
        End Try
    End Sub
    Public Sub Validar_Medicamentos(ByRef id As String, ByRef PExcluir As String)
        Dim sSQL As String
        sSQL = "ERRORES_EN_MEDICAMENTOS"
        Try
            Using cn As New MySqlConnection(conexion)
                cn.Open()
                Using cmd As New MySqlCommand(sSQL, cn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandTimeout = 900000000
                    cmd.Parameters.Add("USession", MySqlDbType.VarChar).Value = id
                    cmd.Parameters.Add("Excluir", MySqlDbType.Float).Value = PExcluir
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As MySqlException
        End Try
    End Sub
    Public Sub Validar_Nacimientos(ByRef id As String, ByRef PExcluir As String)
        Dim sSQL As String
        sSQL = "ERRORES_EN_RECIEN_NACIDOS"
        Try
            Using cn As New MySqlConnection(conexion)
                cn.Open()
                Using cmd As New MySqlCommand(sSQL, cn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandTimeout = 900000000
                    cmd.Parameters.Add("USession", MySqlDbType.VarChar).Value = id
                    cmd.Parameters.Add("Excluir", MySqlDbType.Float).Value = PExcluir
                    cmd.ExecuteNonQuery()
                    cn.Close()
                End Using
            End Using
        Catch ex As MySqlException
            MsgBox(ex.Message, , sSQL)
        End Try
    End Sub
    Public Sub Validar_Otros_servicios(ByRef id As String, ByRef PExcluir As String, ByRef porce As String)
        Dim sSQL As String
        sSQL = "ERRORES_EN_OTROS_SERVICIOS"
        Try
            Using cn As New MySqlConnection(conexion)
                cn.Open()
                Using cmd As New MySqlCommand(sSQL, cn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandTimeout = 900000000
                    cmd.Parameters.Add("USession", MySqlDbType.VarChar).Value = id
                    cmd.Parameters.Add("Excluir", MySqlDbType.Float).Value = PExcluir
                    cmd.Parameters.Add("porce", MySqlDbType.Float).Value = CInt(porce)
                    cmd.ExecuteNonQuery()
                    cn.Close()
                End Using
            End Using
        Catch ex As MySqlException
        End Try
    End Sub

    'Public Sub Validar_Otros_serviciostari(ByRef id As String, ByRef Porcentajetxt As String, ByRef CbTipoTarifa As String)

    '    Dim SSQL As String = ""


    '    porce = CInt(Porcentajetxt)
    '    tari = CbTipoTarifa
    '    Dim iResultado As Integer
    '    Try
    '        Using conn As New MySqlConnection(conexion)
    '            conn.Open()
    '            SSQL = "SET AUTOCOMMIT = 0;  INSERT INTO error_at_ (NUM_FACTURA, TIPO_IDENTIFI, NUM_IDENTIFI, TIPO_SERV, CODIGO_SERV, NOMBRE_SERV, CANT , VALOR_UNITARIO, DESCRIPCION_DEL_ERROR , ERROR1 , Usuario) SELECT a.Campo1 , a.Campo3 , a.Campo4 , a.Campo6 , a.Campo7 , a.Campo8 , a.Campo9 ,a.Campo10 , CONCAT('DIFERENCIA DE TARIFA - ','TARIFA :',t.VALOR,' - PORCENTAJE ('," & porce & ",'%) : ',ROUND(t.VALOR *(" & porce & " / 100), -2), ' - DIFERENCIA :',(a.Campo10-ROUND(t.VALOR *(" & porce & " / 100), -2))) AS DESCRIP, a.Campo9*(a.Campo10-ROUND(t.VALOR *(" & porce & " / 100), -2)) AS PORC,'" & id & "' FROM at01 a INNER JOIN tarifas_1 t ON t.CÓDIGO=a.Campo7 AND t.AÑO = YEAR(a.FECHA) WHERE t.MANUAL = 'cups' AND ROUND(t.VALOR *(" & porce & " / 100), 1)<>a.Campo10 AND (a.Campo10-ROUND(t.VALOR *(" & porce & " / 100), 1))>0 AND a.Usuario='" & id & "';  COMMIT;"
    '            Dim cmd As New MySqlCommand(SSQL, conn)
    '            cmd.CommandType = CommandType.Text
    '            cmd.CommandTimeout = 9000000
    '            iResultado = cmd.ExecuteNonQuery() ' ejecutar comando 
    '            conn.Close()
    '        End Using
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Exclamation, "VALIDANDO TARIFAS EN OTROS SERVICIOS ")
    '    End Try
    'End Sub
    Public Sub Validar_Procedimientos(ByRef id As String, ByRef PExcluir As String, ByRef porc As String)
        Dim sSQL As String
        sSQL = "ERRORES_EN_PROCEDIMIENTOS"
        Try
            Using cn As New MySqlConnection(conexion)
                cn.Open()

                Using cmd As New MySqlCommand(sSQL, cn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandTimeout = 900000000
                    cmd.Parameters.Add("USession", MySqlDbType.VarChar).Value = id
                    cmd.Parameters.Add("Excluir", MySqlDbType.Float).Value = PExcluir
                    cmd.Parameters.Add("porce", MySqlDbType.Float).Value = CInt(porc)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As MySqlException
            MsgBox(ex.Message, , sSQL)
        End Try
    End Sub

    'Public Sub Validar_Procedimientostari(ByRef id As String, ByRef Porcentajetxt As String, ByRef CbTipoTarifa As String)
    '    Dim cs As String, SSQL As String
    '    por1 = CInt(Porcentajetxt)
    '    cs = ""
    '    SSQL = ""
    '    porce = CInt(Porcentajetxt)
    '    tari = CbTipoTarifa
    '    Try
    '        Using conn As New MySqlConnection(conexion)
    '            conn.Open()
    '            SSQL = "SET AUTOCOMMIT = 0; INSERT INTO error_ap_ (TIPO_IDENTIFI, NUM_IDENTIFI, NUM_FACTURA, FECHA_PROC, CODIGO_PROC, DESCRIPCION_DEL_ERROR , ERROR1 ,Usuario) SELECT ap.Campo3 , ap.Campo4 , ap.Campo1 , ap.Campo5 , ap.Campo7 , 'DIFERENCIA DE TARIFA' AS DESCRIP, ap.Campo15-ROUND(t.VALOR *(" & porce & " / 100), -2) PORC,'" & id & "' FROM ap INNER JOIN tarifas_1 t ON t.CÓDIGO=ap.Campo7 AND t.AÑO = YEAR(ap.Campo5) WHERE t.MANUAL = 'cups' AND ROUND(t.VALOR *(" & porce & " / 100), 1)<>ap.Campo15 AND (ap.Campo15-ROUND(t.VALOR *(" & porce & " / 100), 1))>0 AND ap.Usuario='" & id & "';  COMMIT;"
    '            Dim cmd As New MySqlCommand(SSQL, conn)
    '            cmd.CommandTimeout = 9000000
    '            Dim iResultado As Integer
    '            iResultado = cmd.ExecuteNonQuery() ' ejecutar comando 
    '        End Using
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Exclamation, "VALIDANDO TARIFAS EN PROCEDIMIENTOS ")
    '    End Try
    'End Sub
    Public Sub Validar_Urgencias(ByRef id As String, ByRef PExcluir As String)
        Dim sSQL As String
        sSQL = "ERRORES_EN_URGENCIAS"
        Try
            Using conn As New MySqlConnection(conexion)
                conn.Open()
                Using cmd As New MySqlCommand(sSQL, conn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandTimeout = 900000000
                    cmd.Parameters.Add("USession", MySqlDbType.VarChar).Value = id
                    cmd.Parameters.Add("Excluir", MySqlDbType.Float).Value = PExcluir
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As MySqlException

        End Try
    End Sub
    Public Sub Validar_Usuarios(ByRef id As String, ByRef PExcluir As String)
        Dim sSQL As String
        sSQL = "ERRORES_EN_USUARIOS"
        Try
            Using cn As New MySqlConnection(conexion)
                cn.Open()
                Using cmd As New MySqlCommand(sSQL, cn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandTimeout = 900000000
                    cmd.Parameters.Add("USession", MySqlDbType.VarChar).Value = id
                    cmd.Parameters.Add("Excluir", MySqlDbType.Float).Value = PExcluir
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As MySqlException

        End Try
    End Sub
    Public Sub Validar_Transaccion(ByRef id As String, ByRef PExcluir As String)
        Dim sSQL As String
        sSQL = "ERRORES_EN_TRANSACCIONES"
        Try
            Using cn As New MySqlConnection(conexion)
                cn.Open()
                Using cmd As New MySqlCommand(sSQL, cn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandTimeout = 900000000
                    cmd.Parameters.Add("USession", MySqlDbType.VarChar).Value = id
                    cmd.Parameters.Add("Excluir", MySqlDbType.Float).Value = PExcluir
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As MySqlException

        End Try
    End Sub
    Public Sub TotalFacturado(ByRef IdUsuariA As String, ByRef PExcluir As String)
        Dim query As String = "TotalFacturado"
        Try
            Using conn As New MySqlConnection(conexion)
                conn.Open()
                Using cmd As New MySqlCommand(query, conn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandTimeout = 900000000
                    cmd.Parameters.Add("USession", MySqlDbType.VarChar).Value = IdUsuariA
                    cmd.Parameters.Add("Excluir", MySqlDbType.Float).Value = PExcluir
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As Exception

        End Try
    End Sub
End Class
