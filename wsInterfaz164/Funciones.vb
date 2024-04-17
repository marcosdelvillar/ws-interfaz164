Imports System.Data
Imports System.Text
Imports System.Security.Cryptography
Imports System.IO
Imports System.Xml
Imports System.Data.SqlClient
Imports Ionic.Zip
Imports System.Xml.XPath

Public Module funciones
    Public PalabraSecreta As String = "$@1Q2X3_esther##"


    Public Function EncryptString(ByVal InputString As String, ByVal SecretKey As String, Optional ByVal CyphMode As CipherMode = CipherMode.ECB) As String
        Try
            Dim Des As New TripleDESCryptoServiceProvider
            'Put the string into a byte array
            Dim InputbyteArray() As Byte = Encoding.UTF8.GetBytes(InputString)
            'Create the crypto objects, with the key, as passed in
            Dim hashMD5 As New MD5CryptoServiceProvider
            Des.Key = hashMD5.ComputeHash(ASCIIEncoding.ASCII.GetBytes(SecretKey))
            Des.Mode = CyphMode
            Dim ms As MemoryStream = New MemoryStream
            Dim cs As CryptoStream = New CryptoStream(ms, Des.CreateEncryptor(),
            CryptoStreamMode.Write)
            'Write the byte array into the crypto stream
            '(It will end up in the memory stream)
            cs.Write(InputbyteArray, 0, InputbyteArray.Length)
            cs.FlushFinalBlock()
            'Get the data back from the memory stream, and into a string
            Dim ret As StringBuilder = New StringBuilder
            Dim b() As Byte = ms.ToArray
            ms.Close()
            Dim I As Integer
            For I = 0 To UBound(b)
                'Format as hex
                ret.AppendFormat("{0:X2}", b(I))
            Next

            Return ret.ToString()
        Catch ex As System.Security.Cryptography.CryptographicException
            'ExceptionManager.Publish(ex)
            Return ""
        End Try

    End Function

    Public Function DecryptString(ByVal InputString As String, ByVal SecretKey As String, Optional ByVal CyphMode As CipherMode = CipherMode.ECB) As String
        Try

            If InputString = String.Empty Then
                Return ""
            Else
                Dim Des As New TripleDESCryptoServiceProvider
                'Put the string into a byte array
                Dim InputbyteArray(CType(InputString.Length / 2 - 1, Integer)) As Byte '= Encoding.UTF8.GetBytes(InputString)
                'Create the crypto objects, with the key, as passed in
                Dim hashMD5 As New MD5CryptoServiceProvider

                Des.Key = hashMD5.ComputeHash(ASCIIEncoding.ASCII.GetBytes(SecretKey))
                Des.Mode = CyphMode
                'Put the input string into the byte array

                Dim X As Integer

                For X = 0 To InputbyteArray.Length - 1
                    Dim IJ As Int32 = (Convert.ToInt32(InputString.Substring(X * 2, 2), 16))
                    Dim BT As New System.ComponentModel.ByteConverter 'ComponentModel.ByteConverter
                    InputbyteArray(X) = New Byte
                    InputbyteArray(X) = CType(BT.ConvertTo(IJ, GetType(Byte)), Byte)
                Next

                Dim ms As MemoryStream = New MemoryStream
                Dim cs As CryptoStream = New CryptoStream(ms, Des.CreateDecryptor(),
                CryptoStreamMode.Write)

                'Flush the data through the crypto stream into the memory stream
                cs.Write(InputbyteArray, 0, InputbyteArray.Length)
                cs.FlushFinalBlock()

                '//Get the decrypted data back from the memory stream
                Dim ret As StringBuilder = New StringBuilder
                Dim B() As Byte = ms.ToArray

                ms.Close()

                Dim I As Integer

                For I = 0 To UBound(B)
                    ret.Append(Chr(B(I)))
                Next

                Return ret.ToString()
            End If

        Catch ex As Exception
            Return "Error: " & ex.Message.ToString
        End Try
    End Function

    Public Function EjecutaDatasetSP(ByVal StoredProcedure As String,
                                    Optional ByVal Parametros As Hashtable = Nothing) As DataSet
        Dim cmdCommand As Data.SqlClient.SqlCommand
        Dim con As New SqlClient.SqlConnection
        Try
            Dim strConnectionString As String = ConfigurationManager.ConnectionStrings("ServidorConnWS2").ConnectionString
            con.ConnectionString = strConnectionString
            If con.State <> ConnectionState.Open Then
                con.Open()
            End If
            cmdCommand = New SqlCommand
            cmdCommand.CommandText = StoredProcedure
            cmdCommand.CommandType = CommandType.StoredProcedure
            cmdCommand.Connection = con

            If Not Parametros Is Nothing Then
                Dim Col As String
                For Each KeyHash As Object In Parametros.Keys
                    Col = KeyHash.ToString
                    cmdCommand.Parameters.Add(New SqlClient.SqlParameter(Col, Parametros(Col)))
                Next
            End If
            Dim daEjecutaDs As SqlDataAdapter = New SqlDataAdapter
            Dim dsEjecutaDs As DataSet = New DataSet
            daEjecutaDs.SelectCommand = cmdCommand
            daEjecutaDs.Fill(dsEjecutaDs)
            con.Close()
            Return dsEjecutaDs
        Catch ex As Exception
            Dim strMensaje As String = ex.Message
            con.Close()
        End Try
        Return Nothing
    End Function

    Public Function EjecutaDataTableSP(ByVal StoredProcedure As String,
                                    Optional ByVal Parametros As Hashtable = Nothing) As Data.DataTable
        Dim cmdCommand As Data.SqlClient.SqlCommand
        Dim con As New SqlClient.SqlConnection
        Try
            Dim strConnectionString As String = ConfigurationManager.ConnectionStrings("ServidorConnWS2").ConnectionString
            con.ConnectionString = strConnectionString
            If con.State <> ConnectionState.Open Then
                con.Open()
            End If
            cmdCommand = New SqlCommand
            cmdCommand.CommandText = StoredProcedure
            cmdCommand.CommandType = CommandType.StoredProcedure
            cmdCommand.Connection = con

            If Not Parametros Is Nothing Then
                Dim Col As String
                For Each KeyHash As Object In Parametros.Keys
                    Col = KeyHash.ToString
                    cmdCommand.Parameters.Add(New SqlClient.SqlParameter(Col, Parametros(Col)))
                Next
            End If

            Dim daEjecutaDt As SqlDataAdapter = New SqlDataAdapter
            Dim dtEjecutaDt As Data.DataTable = New Data.DataTable
            daEjecutaDt.SelectCommand = cmdCommand
            daEjecutaDt.Fill(dtEjecutaDt)
            con.Close()
            Return dtEjecutaDt

        Catch ex As Exception
            Dim strMensaje As String = ex.Message
            con.Close()
        End Try
        Return Nothing
    End Function

    Public Function EjecutaSP(ByVal StoredProcedure As String, Optional ByVal Parametros As Hashtable = Nothing) As String
        Dim cmdCommand As Data.SqlClient.SqlCommand
        Dim con As New SqlClient.SqlConnection
        Try
            Dim strConnectionString As String = ConfigurationManager.ConnectionStrings("ServidorConnWS2").ConnectionString
            con.ConnectionString = strConnectionString
            If con.State <> ConnectionState.Open Then
                con.Open()
            End If
            cmdCommand = New SqlCommand
            cmdCommand.CommandText = StoredProcedure
            cmdCommand.CommandType = CommandType.StoredProcedure
            cmdCommand.Connection = con
            If Not Parametros Is Nothing Then
                Dim Col As String
                For Each KeyHash As Object In Parametros.Keys
                    Col = KeyHash.ToString
                    cmdCommand.Parameters.Add(New SqlClient.SqlParameter(Col, Parametros(Col)))
                Next
            End If
            cmdCommand.ExecuteNonQuery()
            con.Close()
            Return "TODOOK"
        Catch ex As Exception
            con.Close()
            Return ex.Message.ToString

        End Try
        Return False
    End Function

    Public Function EjecutaSP(ByVal StoredProcedure As String, ByVal Parametros As Hashtable, ByVal NameParamIdentity As String) As Object
        Dim cmdCommand As Data.SqlClient.SqlCommand
        Dim con As New SqlClient.SqlConnection
        Try
            Dim strConnectionString As String = ConfigurationManager.ConnectionStrings("ServidorConnWS2").ConnectionString
            con.ConnectionString = strConnectionString
            If con.State <> ConnectionState.Open Then
                con.Open()
            End If
            cmdCommand = New SqlCommand
            cmdCommand.CommandText = StoredProcedure
            cmdCommand.CommandType = CommandType.StoredProcedure
            cmdCommand.Connection = con
            If Not Parametros Is Nothing Then
                Dim Col As String
                For Each KeyHash As Object In Parametros.Keys
                    Col = KeyHash.ToString
                    cmdCommand.Parameters.Add(New SqlClient.SqlParameter(Col, Parametros(Col)))
                Next
            End If

            cmdCommand.Parameters.Add(New SqlClient.SqlParameter(NameParamIdentity, SqlDbType.Int, 4))
            cmdCommand.Parameters(NameParamIdentity).Direction = ParameterDirection.Output

            cmdCommand.ExecuteNonQuery()

            Dim ultimo As Object
            ultimo = CType(cmdCommand.Parameters(NameParamIdentity).Value, Object)

            con.Close()

            Return ultimo
        Catch ex As Exception
            Dim strMensaje As String = ex.Message
            con.Close()
            Return Nothing
        End Try
    End Function

    Public Function EjecutaDatasetQry(ByVal strQry As String) As DataSet
        Dim con As New SqlClient.SqlConnection
        Try
            Dim strConnectionString As String = ConfigurationManager.ConnectionStrings("ServidorConnWS2").ConnectionString
            con.ConnectionString = strConnectionString
            If con.State <> ConnectionState.Open Then
                con.Open()
            End If
            Dim daEjecutaDs As SqlDataAdapter = New SqlDataAdapter(strQry, con)
            Dim dsEjecutaDs As DataSet = New DataSet
            daEjecutaDs.Fill(dsEjecutaDs)
            con.Close()
            Return dsEjecutaDs

        Catch ex As Exception
            Dim strMensaje As String = ex.Message
            con.Close()
        End Try
        Return Nothing
    End Function

    Public Function EjecutaDataTableQry(ByVal strQry As String) As DataTable
        Dim con As New SqlClient.SqlConnection
        Try
            Dim strConnectionString As String = ConfigurationManager.ConnectionStrings("ServidorConnWS2").ConnectionString
            con.ConnectionString = strConnectionString
            If con.State <> ConnectionState.Open Then
                con.Open()
            End If
            Dim daEjecutaDt As SqlDataAdapter = New SqlDataAdapter(strQry, con)
            Dim dtEjecutaDt As DataTable = New DataTable
            daEjecutaDt.Fill(dtEjecutaDt)
            con.Close()
            Return dtEjecutaDt
        Catch ex As Exception
            Dim strMensaje As String = ex.Message
            con.Close()
        End Try
        Return Nothing
    End Function

    Public Function EjecutaQuery(ByVal strQry As String) As Boolean
        Dim cmdCommand As New Data.SqlClient.SqlCommand
        Dim con As New SqlClient.SqlConnection
        Try
            Dim strConnectionString As String = ConfigurationManager.ConnectionStrings("ServidorConnWS2").ConnectionString
            con.ConnectionString = strConnectionString
            If con.State <> ConnectionState.Open Then
                con.Open()
            End If
            cmdCommand.CommandText = strQry
            cmdCommand.CommandType = CommandType.Text
            cmdCommand.Connection = con
            cmdCommand.ExecuteNonQuery()
            con.Close()
            Return True
        Catch ex As Exception
            Dim strMensaje As String = ex.Message
            Return False
        End Try
    End Function

    Public Function ExecuteSP_Int32(ByVal spName As String, ByRef sqlParameters() As SqlParameter) As Int32
        'Ejecuta un stored Procedure y devuelve el resultado formateado como int32
        Dim cmd As New SqlClient.SqlCommand
        Dim sConn As String = ConfigurationManager.ConnectionStrings("ServidorConnWS2").ConnectionString
        Dim var_connection As SqlConnection = New SqlConnection(sConn)

        Try
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = spName
            If Not sqlParameters Is Nothing Then
                Dim sqlP As SqlClient.SqlParameter
                For Each sqlP In sqlParameters
                    cmd.Parameters.Add(sqlP)
                Next
            End If
            cmd.Connection = var_connection
            cmd.Connection.Open()
            Return CType(cmd.ExecuteScalar, Int32)
        Catch ex As Exception
            cmd.Connection.Close()
            cmd.Dispose()
            cmd = Nothing
            'Throw New Exception("INT-002 The last command was unsuccessful because the following reason: " & ex.Message)
            Return -1
        Finally
            cmd.Connection.Close()
            cmd.Dispose()
            cmd = Nothing
        End Try
    End Function

    Public Function EjecutaDatasetSP_OUT(ByVal StoredProcedure As String,
                                         ByVal sXML As String,
                                         ByRef Prefijo As String,
                                         ByRef Factura As Integer,
                                         ByRef sError As String) As String

        Dim cmdCommand As Data.SqlClient.SqlCommand
        Dim con As New SqlClient.SqlConnection
        Try
            Dim strConnectionString As String = ConfigurationManager.ConnectionStrings("ServidorConnWS2").ConnectionString
            con.ConnectionString = strConnectionString
            If con.State <> ConnectionState.Open Then
                con.Open()
            End If

            cmdCommand = New SqlCommand
            cmdCommand.CommandText = StoredProcedure
            cmdCommand.CommandType = CommandType.StoredProcedure
            cmdCommand.Connection = con

            cmdCommand.Parameters.Add("@dta", SqlDbType.Xml)
            cmdCommand.Parameters.Item("@dta").Direction = ParameterDirection.Input
            cmdCommand.Parameters.Item("@dta").Value = sXML

            cmdCommand.Parameters.Add("@fact", SqlDbType.Int)
            cmdCommand.Parameters.Item("@fact").Direction = ParameterDirection.Output

            cmdCommand.Parameters.Add("@prefijo", SqlDbType.VarChar, 6)
            cmdCommand.Parameters.Item("@prefijo").Direction = ParameterDirection.Output

            cmdCommand.Parameters.Add("@err", SqlDbType.VarChar, 300)
            cmdCommand.Parameters.Item("@err").Direction = ParameterDirection.Output

            cmdCommand.CommandTimeout = 50000

            ' daEjecutaDs.SelectCommand.CommandTimeout = 40000

            cmdCommand.ExecuteNonQuery()

            Factura = cmdCommand.Parameters.Item("@fact").Value
            Prefijo = cmdCommand.Parameters.Item("@prefijo").Value
            sError = cmdCommand.Parameters.Item("@err").Value

            con.Close()

            Return "TODOOK"
        Catch ex As Exception
            con.Close()
            Return ex.Message.ToString

        End Try
        Return Nothing
    End Function

    Public Function FillDataSet(ByVal StoredProcedure As String,
                                ByRef sRespuesta As String,
                                Optional ByVal Parametros As Hashtable = Nothing) As DataSet

        Dim cmdCommand As Data.SqlClient.SqlCommand
        Dim con As New SqlClient.SqlConnection
        Try
            Dim strConnectionString As String = ConfigurationManager.ConnectionStrings("ConnWSSICOBONv2").ConnectionString
            con.ConnectionString = strConnectionString
            If con.State <> ConnectionState.Open Then
                con.Open()
            End If
            cmdCommand = New SqlCommand
            cmdCommand.CommandText = StoredProcedure
            cmdCommand.CommandType = CommandType.StoredProcedure
            cmdCommand.Connection = con

            If Not Parametros Is Nothing Then
                Dim Col As String
                For Each KeyHash As Object In Parametros.Keys
                    Col = KeyHash.ToString
                    cmdCommand.Parameters.Add(New SqlClient.SqlParameter(Col, Parametros(Col)))
                Next
            End If
            Dim daEjecutaDs As SqlDataAdapter = New SqlDataAdapter
            Dim dsEjecutaDs As DataSet = New DataSet
            daEjecutaDs.SelectCommand = cmdCommand
            daEjecutaDs.Fill(dsEjecutaDs)
            con.Close()

            sRespuesta = "TODOOK"
            Return dsEjecutaDs
        Catch ex As Exception
            sRespuesta = ex.Message
            If con.State = ConnectionState.Open Then con.Close()
            Return Nothing
        End Try

    End Function

    Public Function Escribir_Archivo(ByVal sCONTENIDO As String, ByVal sDESTINO As String) As String

        Dim strArchivo As String = sDESTINO

        Try
            If FileIO.FileSystem.FileExists(strArchivo) Then FileIO.FileSystem.DeleteFile(strArchivo)

            FileOpen(2, strArchivo, OpenMode.Append, OpenAccess.Write, OpenShare.LockWrite)
            PrintLine(2, sCONTENIDO)
            FileClose(2)

            Return "TODOOK"
        Catch ex As Exception
            Return "Error al escribir archivo: " & ex.Message.ToString
        End Try

    End Function

    Public Function EscribirBitacora(ByVal sCONTENIDO As String, ByVal sDESTINO As String) As String

        Dim strArchivo As String = sDESTINO

        Try
            'If FileIO.FileSystem.FileExists(strArchivo) Then FileIO.FileSystem.DeleteFile(strArchivo)

            FileOpen(2, strArchivo, OpenMode.Append, OpenAccess.Write, OpenShare.LockWrite)
            PrintLine(2, Now.ToString("dd/MMM/yy HH:mm:ss").ToUpper & " " & sCONTENIDO)
            FileClose(2)

            Return "TODOOK"
        Catch ex As Exception
            Return "Error al escribir archivo: " & ex.Message.ToString
        End Try

    End Function

    Public Function LeerArchivo(ByVal FileName As String) As String
        Dim oRead As New System.IO.StreamReader
        Dim LineIn As New StringBuilder

        Try
            'SE REVISA QUE EXISTA EL ARCHIVO 
            If Not File.Exists(FileName) Then
                Return "Error, El archivo de configuración no existe en la ruta: " & FileName
            End If

            'SE REALIZA LA LECTURA DEL ARCHIVO
            oRead = File.OpenText(FileName)

            'SE RECORRE TODO EL STREAM DEL ARCHIVO
            While Not oRead.EndOfStream
                'SE ALMACENA EL CONTENIDO DE LA LINEA EN LA VARIABLE
                LineIn.Append(oRead.ReadLine())
            End While

            'SE CIERRA EL STREAM DEL ARCHIVO
            oRead.Close()

            Return LineIn.ToString
        Catch ex As Exception
            oRead.Close()
            Return "Error en 'LeerArchivo': " & ex.Message.ToString
        End Try

    End Function

    Public Sub LimpiarCadena(ByRef CadEntrada As String)
        Try

            CadEntrada = CadEntrada.Replace("Á", "A")
            CadEntrada = CadEntrada.Replace("É", "E")
            CadEntrada = CadEntrada.Replace("Í", "I")
            CadEntrada = CadEntrada.Replace("Ó", "O")
            CadEntrada = CadEntrada.Replace("Ú", "U")
            CadEntrada = CadEntrada.Replace("Ñ", "N")

            CadEntrada = CadEntrada.Replace("á", "a")
            CadEntrada = CadEntrada.Replace("é", "e")
            CadEntrada = CadEntrada.Replace("í", "i")
            CadEntrada = CadEntrada.Replace("ó", "o")
            CadEntrada = CadEntrada.Replace("ú", "u")
            CadEntrada = CadEntrada.Replace("ñ", "n")

        Catch ex As Exception
            CadEntrada = "[ERROR 01A]"
        End Try

    End Sub

    Public Function ZipDirectory(ByVal sDir As String, ByVal ZipFilePath As String) As String
        Try


            Using zip As ZipFile = New ZipFile
                zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestCompression
                'zip.Encryption = EncryptionAlgorithm.WinZipAes256
                'zip.Password = "_ether$#12"

                Dim di As New IO.DirectoryInfo(sDir)
                Dim diar1 As IO.FileInfo() = di.GetFiles()
                Dim dra As IO.FileInfo

                'list the names of all files in the specified directory
                For Each dra In diar1
                    zip.AddFile(dra.FullName, "")
                Next

                zip.Save(ZipFilePath)

            End Using

            Return "TODOOK"
        Catch ex1 As Exception
            Return "ERROR ZIP01: " & ex1.Message.ToString
        End Try
    End Function

    Public Function CrearDataSetFromXMLString(ByVal sXML As String) As DataSet
        Dim ds As New DataSet
        Dim oDoc As XmlDocument

        Try
            ds = New DataSet
            oDoc = New XmlDocument
            oDoc.LoadXml(sXML)

            ds.ReadXml(New XmlNodeReader(oDoc))

            Return ds
        Catch ex As Exception
            MsgBox("Error en 'CrearDataSetFromXMLString': " & ex.Message.ToString)
            Return Nothing
        End Try
    End Function

    Public Function EnviarMail(ByVal strSubject As String,
                                 ByVal strBody As String,
                                 ByVal strTO As String,
                                 ByVal strCCO As String,
                                 ByVal strAttach As String) As String

        Dim configurationAppSettings As New System.Configuration.AppSettingsReader

        Try

            Dim Partes() As String
            Dim miEmail As New System.Net.Mail.MailMessage

            Try
                With miEmail
                    Dim mAddress As New System.Net.Mail.MailAddress(CType(configurationAppSettings.GetValue("MailFrom", GetType(System.String)), String), CType(configurationAppSettings.GetValue("MailShowName", GetType(System.String)), String))

                    .From = mAddress

                    Partes = strTO.Split(";")
                    For Each cadena As String In Partes
                        If cadena.Trim.Length > 0 Then .To.Add(cadena)
                    Next

                    .Subject = strSubject
                    .Body = strBody
                    .IsBodyHtml = True

                    If strAttach.Trim.Length > 0 Then
                        For Each cadena As String In strAttach.Split(";")
                            'Si existe el archivo, anexarlo
                            If System.IO.File.Exists(cadena) Then
                                Dim att As New System.Net.Mail.Attachment(cadena.Trim)
                                .Attachments.Add(att)
                            End If
                        Next
                    End If

                    Partes = strCCO.Split(";")
                    For Each cadena As String In Partes
                        If cadena.Trim.Length > 0 Then .Bcc.Add(cadena)
                    Next


                End With

                Dim LogoPath As String = CType(configurationAppSettings.GetValue("Path", GetType(System.String)), String) & "FLPLogo.png"

                If System.IO.File.Exists(LogoPath) Then
                    Dim VistaAlterna As System.Net.Mail.AlternateView
                    VistaAlterna = System.Net.Mail.AlternateView.CreateAlternateViewFromString(strBody, Nothing, "text/html")

                    Dim Logo As New System.Net.Mail.LinkedResource(LogoPath)
                    Logo.ContentId = "logo"
                    VistaAlterna.LinkedResources.Add(Logo)

                    miEmail.AlternateViews.Add(VistaAlterna)
                End If


                Dim miSMTP As New System.Net.Mail.SmtpClient
                miSMTP.UseDefaultCredentials = False
                miSMTP.Credentials = New System.Net.NetworkCredential(CType(configurationAppSettings.GetValue("MailUser", GetType(System.String)), String), CType(configurationAppSettings.GetValue("MailPWD", GetType(System.String)), String))
                miSMTP.Port = CType(configurationAppSettings.GetValue("MailPort", GetType(System.String)), String)
                miSMTP.Host = CType(configurationAppSettings.GetValue("MailHost", GetType(System.String)), String)
                miSMTP.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network
                miSMTP.EnableSsl = True
                miSMTP.Send(miEmail) 'Enviar Email

                miSMTP = Nothing
            Catch eX As Exception
                miEmail.Dispose()
                miEmail = Nothing
                Return "ERROR 02 al enviar Mail: " & eX.Message.ToString
            End Try

            miEmail.Dispose()
            miEmail = Nothing
            Return "TODOOK"
        Catch ex As Exception
            Return "ERROR 01: " & ex.Message.ToString
        End Try

    End Function

    Public Function LeerArchivo1Time(ByVal RutaArchivo As String) As String
        Try
            Dim sContenido As String

            sContenido = IO.File.ReadAllText(RutaArchivo, System.Text.Encoding.UTF8)

            Return sContenido
        Catch ex As Exception
            Return "Error: " & ex.Message.ToString
        End Try
    End Function

    Public Function saveDStoFile(ByVal ds As DataSet,
                               ByVal FileName As String) As String

        Dim sb As New StringBuilder
        Const FIELDSEPARATOR As String = "" & vbTab & ""
        Const ROWSEPARATOR As String = "" & vbLf & ""

        Try
            For iTabla As Integer = 0 To ds.Tables.Count - 1
                sb.Append("Tabla #" & iTabla.ToString & ds.Tables(iTabla).TableName.ToString)
                sb.Append(ROWSEPARATOR)
                sb.Append("Columnas")
                sb.Append(ROWSEPARATOR)
                For iColumna As Integer = 0 To ds.Tables(iTabla).Columns.Count - 1
                    sb.Append(ds.Tables(iTabla).Columns(iColumna).ColumnName).Append(FIELDSEPARATOR)
                Next
                sb.Append(ROWSEPARATOR)
                For iFila As Integer = 0 To ds.Tables(iTabla).Rows.Count - 1
                    For iColumna As Integer = 0 To ds.Tables(iTabla).Columns.Count - 1
                        If IsDate(ds.Tables(iTabla).Rows(iFila)(iColumna).ToString) Then
                            sb.Append(Format(CDate(ds.Tables(iTabla).Rows(iFila)(iColumna).ToString), "dd/MMM/yyyy HH:mm:ss")).Append(FIELDSEPARATOR)
                        Else
                            sb.Append(ds.Tables(iTabla).Rows(iFila)(iColumna).ToString).Append(FIELDSEPARATOR)
                        End If

                    Next
                    sb.Append(ROWSEPARATOR)
                Next

            Next

            'Call Escribir_Archivo(sb.ToString, FileName)
            Dim sw As New StreamWriter(FileName)
            sw.Write(sb.ToString())
            sw.Close()
            Return "TODOOK"
        Catch ex As Exception
            Call Escribir_Archivo(ex.Message.ToString, FileName)
            Return ex.Message.ToString
        End Try
    End Function


#Region "Funciones para el RFC"
    Public Function RFCFiltraAcentos(ByVal strTexto As String) As String

        'Esta rutina elimina los acentos y convierte el nombre
        'a mayusculas.

        strTexto = strTexto.ToUpper
        strTexto = Replace(strTexto, "Á", "A")
        strTexto = Replace(strTexto, "É", "E")
        strTexto = Replace(strTexto, "Í", "I")
        strTexto = Replace(strTexto, "Ó", "O")
        strTexto = Replace(strTexto, "Ú", "U")

        Return strTexto

    End Function

    Public Sub RFCFiltraNombres(ByRef strNombre As String,
                                ByRef strPaterno As String,
                                ByRef strMaterno As String)

        'Esta rutina elimina palabras sobrantes para el
        'calculo del RFC de los tres nombres.

        Dim strArPalabras()
        Dim i

        'Inicializa el arreglo con las palabras que no queremos.
        'strArPalabras = Array(".", ",", "DE ", "DEL ", "LA ", "LOS ", "LAS ", "Y ", "MC ", "MAC ", "VON ", "VAN ")
        'ASP no tiene la función bien definida ARRAY de Visual Basic
        ReDim Preserve strArPalabras(1) : strArPalabras(1) = "."
        ReDim Preserve strArPalabras(2) : strArPalabras(2) = ","
        ReDim Preserve strArPalabras(3) : strArPalabras(3) = "DE "
        ReDim Preserve strArPalabras(4) : strArPalabras(4) = "DEL "
        ReDim Preserve strArPalabras(5) : strArPalabras(5) = "LA "
        ReDim Preserve strArPalabras(6) : strArPalabras(6) = "LOS "
        ReDim Preserve strArPalabras(7) : strArPalabras(7) = "LAS "
        ReDim Preserve strArPalabras(8) : strArPalabras(8) = "Y "
        ReDim Preserve strArPalabras(9) : strArPalabras(9) = "MC "
        ReDim Preserve strArPalabras(10) : strArPalabras(10) = "MAC "
        ReDim Preserve strArPalabras(11) : strArPalabras(11) = "VON "
        ReDim Preserve strArPalabras(12) : strArPalabras(12) = "VAN "

        'Busca cada palabra en los tres nombre y eliminala
        'se se encuentra.
        For i = LBound(strArPalabras) To UBound(strArPalabras)
            strNombre = Replace(strNombre, strArPalabras(i), "")
            strPaterno = Replace(strPaterno, strArPalabras(i), "")
            strMaterno = Replace(strMaterno, strArPalabras(i), "")
        Next

        'Listo, ahora sigo con el nombre pila, buscando
        'la presencia de Maria o Jose.

        'Inicializa el arreglo con las palabras que
        'queremos eliminar.
        'strArPalabras = Array("JOSE ", "MARIA ", "J ", "MA ")
        ReDim strArPalabras(0)
        ReDim Preserve strArPalabras(1) : strArPalabras(1) = "JOSE "
        ReDim Preserve strArPalabras(2) : strArPalabras(2) = "MARIA "
        ReDim Preserve strArPalabras(3) : strArPalabras(3) = "J "
        ReDim Preserve strArPalabras(4) : strArPalabras(4) = "MA "

        'Haz esto solo si el nombre de pila tiene algun
        'espacio.
        If InStr(strNombre, " ") > 0 Then
            For i = LBound(strArPalabras) To UBound(strArPalabras)
                strNombre = Replace(strNombre, strArPalabras(i), "")
            Next
        End If


        'Por ultimo, elimina doble consonantes de los nombres
        'cuando estas ocurren en las primeras dos letras del
        'nombre.
        Select Case Microsoft.VisualBasic.Left(strNombre, 2)
            Case "CH"
                strNombre = Replace(strNombre, "CH", "C", 1, 1)
            Case "LL"
                strNombre = Replace(strNombre, "LL", "L", 1, 1)
        End Select

        Select Case Microsoft.VisualBasic.Left(strPaterno, 2)
            Case "CH"
                strPaterno = Replace(strPaterno, "CH", "C", 1, 1)
            Case "LL"
                strPaterno = Replace(strPaterno, "LL", "L", 1, 1)
        End Select

        Select Case Microsoft.VisualBasic.Left(strMaterno, 2)
            Case "CH"
                strMaterno = Replace(strMaterno, "CH", "C", 1, 1)
            Case "LL"
                strMaterno = Replace(strMaterno, "LL", "L", 1, 1)
        End Select

    End Sub

    Public Function RFCApellidoCorto(ByVal strNombre As String,
                                    ByVal strPaterno As String,
                                    ByVal strMaterno As String,
                                    ByVal strFecha As String)

        'Eta rutina calcula el RFC tomando en cuenta un
        'apellido paterno de tres o menos letras.

        'RFCApellidoCorto = Left(strPaterno, 1) & Left(strMaterno, 1) & Left(strNombre, 2) & strFecha
        RFCApellidoCorto = Microsoft.VisualBasic.Left(strPaterno, 1) & Microsoft.VisualBasic.Left(strMaterno, 1) & Microsoft.VisualBasic.Left(strNombre, 2) & strFecha


    End Function

    Public Function RFCArmalo(ByVal strNombre As String,
                            ByVal strPaterno As String,
                            ByVal strMaterno As String,
                            ByVal strFecha As String) As String

        'Esta rutina arma el RFC basandose en los tres nombres
        'y la fecha de nacimiento.

        'Dim strArVocales() As Variant
        Dim strVocales As String = ""
        Dim strLetra As String = ""
        Dim strPrimerVocal As String = ""
        Dim i As Integer

        'Inicializa la cadena de vocales.
        strVocales = "AEIOU"

        'Primero consigo la primera vocal del nombre
        'comenzando con la segunda letra.
        For i = 2 To Len(strPaterno)
            If InStr(strVocales, Mid(strPaterno, i, 1)) > 0 Then
                strPrimerVocal = Mid(strPaterno, i, 1)
                Exit For
            End If
        Next

        'RFCArmalo = Left(strPaterno, 1) & strPrimerVocal & Left(strMaterno, 1) & Left(strNombre, 1) & strFecha
        Return Microsoft.VisualBasic.Left(strPaterno, 1) & strPrimerVocal & Microsoft.VisualBasic.Left(strMaterno, 1) & Microsoft.VisualBasic.Left(strNombre, 1) & strFecha

    End Function

    Public Function RFCUnApellido(ByVal strNombre As String,
                                ByVal strPaterno As String,
                                ByVal strMaterno As String,
                                ByVal strFecha As String) As String

        'Esta rutina toma en cuenta casos cuando solo se
        'da un apellido, ya sea el paterno o materno.

        Dim strApellido

        Select Case True
            Case Len(strPaterno) > 0 And Len(strMaterno) = 0
                'Solo hay apellido paterno.
                strApellido = strPaterno.Substring(0, 2) 'Left(strPaterno, 2)
            Case Len(strPaterno) = 0 And Len(strMaterno) > 0
                'Solo hay apellido materno.
                strApellido = strMaterno.Substring(0, 2) 'Left(strMaterno, 2)
            Case Else
                strApellido = strNombre.Substring(0, 2) 'Left(strNombre, 2)
        End Select

        'Ahora arma el RFC.
        'RFCUnApellido = strApellido & Left(strNombre, 2) & strFecha
        Return strApellido & Microsoft.VisualBasic.Left(strNombre, 2) & strFecha
        '

    End Function

    Public Function RFCQuitaProhibidas(ByVal strRFC As String) As String

        'Esta rutina quita cualquiera de las palabras prohibidas,
        'cambiando el ultimo caracter de dicha palabra a X.

        Dim strPalabras

        'Define todas las palabras prohibidas.
        strPalabras = "BUEI*BUEY*CACA*CACO*CAGA*CAGO*CAKA*CAKO*COGE*COJA*"
        strPalabras = strPalabras & "KOGE*KOJO*KAKA*KULO*MAME*MAMO*MEAR*"
        strPalabras = strPalabras & "MEAS*MEON*MION*COJE*COJI*COJO*CULO*"
        strPalabras = strPalabras & "FETO*GUEY*JOTO*KACA*KACO*KAGA*KAGO*"
        strPalabras = strPalabras & "MOCO*MULA*PEDA*PEDO*PENE*PUTA*PUTO*"
        strPalabras = strPalabras & "QULO*RATA*RUIN*"

        'Si alguna de estas se encuentra, cambiala.
        If InStr(strPalabras, strRFC.Substring(0, 4) & "*") > 0 Then 'If InStr(strPalabras, Left(strRFC, 4) & "*") > 0 Then
            'Reemplaza el cuarto caracter del RFC para eliminar
            'l apalabra prohibida.
            'Mid(strRFC, 4, 1) = "X"
            strRFC = Microsoft.VisualBasic.Left(strRFC, 3) & "X" & Mid(strRFC, 5)
        End If

        Return strRFC
    End Function

    Public Function RFCHomoclave(ByVal strNombre As String,
                                ByVal strPaterno As String,
                                ByVal strMaterno As String) As String

        'Esta rutina calcula la homoclave, que es de dos
        'caracteres. El proceso solo toma en cuenta los
        'nombres de la persona.

        Dim strNombreComp As String
        Dim strCharsHc As String
        Dim strChr As String
        Dim i As Integer
        Dim strCadena As String
        Dim intNum1, intNum2 As Integer
        Dim intSum As Integer
        Dim int3 As Integer
        Dim intQuo, intRem As Integer

        'Consigue el nombre completo de la persona.
        strNombreComp = strPaterno & " " & strMaterno & " " & strNombre

        'Inicializa la cadena de caracteres que contiene
        'los caracteres permitidos para la homoclave.
        'Notese la ausencia del numero 0 y la letra o.
        strCharsHc = "123456789ABCDEFGHIJKLMNPQRSTUVWXYZ"

        'Inicializa la cadena con 0 para desplazar todo a
        'la derecha.
        strCadena = "0"

        For i = 1 To Len(strNombreComp)
            strChr = Mid(strNombreComp, i, 1)

            'Convierte la letra a un numero de dos
            'digitos.
            Select Case strChr
                Case " ", "-"
                    strCadena = strCadena & "00"
                Case "Ñ", "Ü"
                    strCadena = strCadena & "10"
                Case "A", "B", "C", "D", "E", "F", "G", "H", "I"
                    strCadena = strCadena & CStr(Asc(strChr) - 54)
                Case "J", "K", "L", "M", "N", "O", "P", "Q", "R"
                    strCadena = strCadena & CStr(Asc(strChr) - 53)
                Case "S", "T", "U", "V", "W", "X", "Y", "Z"
                    strCadena = strCadena & CStr(Asc(strChr) - 51)
                Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                    'Se supone que esta linea nunca se ejecutara, pues
                    'un nombre no usa digitos. Aun asi, como estaba
                    'en el algoritmo original, lo dejo aqui.
                    strCadena = strCadena & Format(strChr, "00")
            End Select

        Next

        'Borra toda la cadena y realiza una operacion matematica
        'en cada uno de los digitos.
        '
        'Por cada digitos se toman dos a la vez y se multiplica
        'este numero por el digito de unidades del mismo numero.
        'Ejemplo:
        '
        ' Si la cadena es 01245
        '
        ' Se comienza con el primer digito, se toman dos y luego
        ' se multiplica por la unidad de ese mismo numero:
        '
        ' Primer digito = 0, los dos: 01
        ' Se multiplica "01" (1) por "1"
        ' Se acumula.
        '
        ' Segundo digito = 1, los dos: 12
        ' Se multiplica "12" (12) por "2"
        '
        ' Tercer digito = 2, los dos: 24
        ' Se multiplica "24" (24) por "4"
        ' --etc.
        For i = 1 To Len(strCadena) - 1
            intNum1 = CInt(Mid(strCadena, i, 2))
            intNum2 = CInt(Mid(strCadena, i + 1, 1))
            intSum = intSum + intNum1 * intNum2
        Next

        'De la suma, solo necesito los ultimos
        'tres digitos. La forma mas facil de hacer
        'esto en convirtiendo el numero a cadena,
        'luego tomando los tres digitos de la derecha.

        'strSum = CStr(intSum)
        'strSum = Right(strSum, 3)
        int3 = CInt(Microsoft.VisualBasic.Right(CStr(intSum), 3))


        intQuo = Int(int3 / 34)
        ' intRem = int3 - intQuo * 34
        intRem = int3 Mod 34
        'La homoclave se consigue usando el
        'cociente y el residuo.

        'Se usa el cociente y residio para
        'buscar las letras del homoclave
        'dentro de la tabla de caracteres
        'permitidos.
        Return Mid(strCharsHc, intQuo + 1, 1) & Mid(strCharsHc, intRem + 1, 1)

    End Function

    Public Function RFCDigitoVerificador(ByVal strRFC As String) As String

        'Esta rutina calcula el digito verificador. El RFC
        'consta de las iniciales, los digitos de la fecha
        'de nacimiento y los dos caracteres de la homoclave.

        Dim strChars As String

        Dim i, intIdx As Integer
        Dim strBuffer As String = ""
        Dim strCh As String
        Dim strDV As String

        Dim intSumas As Integer

        Dim intDV As Integer

        strChars = "0123456789ABCDEFGHIJKLMN&OPQRSTUVWXYZ*"


        'El RFC tiene 12 caracteres:
        ' 4 Letras, 6 digitos y 2 caracteres (homoclave)
        '
        'Barre los 12 caracteres del RFC.

        For i = 1 To Len(strRFC)
            strCh = Mid(strRFC, i, 1)
            strCh = IIf(strCh = " ", "*", strCh)

            intIdx = InStr(strChars, strCh) - 1
            'strBuffer = strBuffer & IIf(intIdx > 0, _
            'Mid(strDigitos, intIdx * 2 - 1, 2), _
            '"00")
            intSumas = intSumas + intIdx * (14 - i)

            'Suplir el Format que no existe en ASP.
            If intIdx < 10 Then strBuffer = strBuffer & "0"
            strBuffer = strBuffer & intIdx

        Next

        If intSumas Mod 11 = 0 Then
            strDV = "0"
        Else
            intDV = 11 - intSumas Mod 11
            If intDV > 9 Then
                strDV = "A"
            Else
                strDV = CStr(intDV)
            End If
        End If

        Return strDV

    End Function
#End Region

#Region "CFDI Bonos"

    Public Function quitarNoValidos(ByVal sCadena As String) As String
        sCadena = sCadena.ToUpper
        sCadena = sCadena.Replace("Á", "A")
        sCadena = sCadena.Replace("É", "E")
        sCadena = sCadena.Replace("Í", "I")
        sCadena = sCadena.Replace("Ó", "O")
        sCadena = sCadena.Replace("Ú", "U")
        sCadena = sCadena.Replace(".", "")
        sCadena = sCadena.Replace(",", "")
        sCadena = sCadena.Replace("  ", " ")

        Return sCadena
    End Function
#End Region

End Module
