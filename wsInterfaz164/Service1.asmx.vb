Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Xml
Imports System.Xml.XPath
Imports Gma.QrCodeNet.Encoding.Windows
Imports Gma.QrCodeNet.Encoding.Windows.Render
Imports Gma.QrCodeNet.Encoding
Imports System.Data

<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class Service1
    Inherits System.Web.Services.WebService

    <WebMethod()>
    Public Function version_01() As String
        Return "1.7"
    End Function

    <WebMethod()>
    Public Function wsGeTicket(ByVal sPrefijo As String,
                            ByVal sNumFactura As String,
                            ByVal sCveSucursal As String,
                            ByVal sPwd As String) As String

        If sPwd.Trim <> "hearMe#$780" Then
            'Return "Error WS01: La contraseña es incorrecta."
        ElseIf sPrefijo.Trim.Length = 0 Then
            Return "Error WS02: El prefijo debe estar vacio caracteres."
        ElseIf Not IsNumeric(sNumFactura) Then
            Return "Error WS03: El número de pedido debe ser numérico."

        End If
        sPrefijo = sPrefijo.ToUpper

        Dim sPath As String = ConfigurationManager.AppSettings("PathInvFactRep").ToString
        sPath &= "FactTicket.rpt"
        Dim sFileName As String = "tck" & sPrefijo & sNumFactura & Format(Now, "_hhmmss") & ".pdf"
        Dim sRutaARchivo As String = ConfigurationManager.AppSettings("WhereToSavePDF").ToString & sFileName
        Dim miReporte As New CrystalDecisions.CrystalReports.Engine.ReportDocument

        Try
            If Not IO.File.Exists(sPath) Then
                Return "El archivo no existe en a ruta: " & sPath
            End If

            miReporte.Load(sPath, CrystalDecisions.Shared.OpenReportMethod.OpenReportByTempCopy)

            miReporte.SetParameterValue("@NumFactura", sNumFactura)
            miReporte.SetParameterValue("@Prefijo", sPrefijo)
            miReporte.SetParameterValue("@CveSucursal", CInt(sCveSucursal))

            miReporte.ExportToDisk(5, sRutaARchivo)

            miReporte.Close()
            miReporte.Dispose()

            Dim sURL As String = ConfigurationManager.AppSettings("LinkDownloadInvfact").ToString & sFileName
            Return sURL

        Catch ex As Exception
            miReporte.Close()
            miReporte.Dispose()
            Return "Error al exportar: " & ex.Message.ToString
        End Try
    End Function

    ''' <summary>
    ''' WS para el envio en PDF de la factura
    ''' </summary>
    ''' <param name="sNoPed">Numero de pedido</param>
    ''' <param name="sPrefijo">Prefijo del pedido</param>
    ''' <param name="sCveEmpresario">Clave del empresario que generó el pedido</param>
    ''' <param name="sFechaAlta">Fecha de alta del pedido dd/MM/aaaa</param>
    ''' <param name="sPwd">Password para conectarse al WS</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()>
    Public Function wsGenerarFactura(ByVal sNoPed As String,
                                    ByVal sPrefijo As String,
                                    ByVal sCveEmpresario As String,
                                    ByVal sFechaAlta As String,
                                    ByVal sPwd As String,
                                    ByVal EnviarCorreo As String) As String

        If sPwd.Trim <> "hearMe#$780" Then
            'Return "Error WS01: La contraseña es incorrecta."
        ElseIf sPrefijo.Trim.Length = 0 Then
            Return "Error WS02: El prefijo debe estar vacio caracteres."
        ElseIf Not IsNumeric(sNoPed) Then
            Return "Error WS03: El número de pedido debe ser numérico."
        ElseIf Not IsDate(sFechaAlta) Then
            Return "Error WS04: La fecha no es valida. Debe estar en el formato dd/MM/aaaa."
        End If

        sPrefijo = sPrefijo.ToUpper

        Dim sRES As String
        Dim HsInfo As Hashtable
        HsInfo = New Hashtable
        HsInfo.Add("@CveEmpresario", CDbl(sCveEmpresario))
        HsInfo.Add("@NumFactura", CDbl(sNoPed))
        HsInfo.Add("@Prefijo", sPrefijo)
        HsInfo.Add("@Fechaalta", sFechaAlta)

        Dim ds As DataSet
        ds = EjecutaDatasetSP("[spTck_SelRevFact]", HsInfo)
        If ds Is Nothing Then
            Return "Error WS05: El SP [spTck_SelRevFact] devolvió nothing."
        ElseIf ds.Tables(0).Rows.Count = 0 Then
            Return "Error WS06: El SP [spTck_SelRevFact] devolvió 0 registros."
        End If

        Dim Fila As DataRow
        Fila = ds.Tables(0).Rows(0)

        If CDbl(Fila("NumFactura")) = 0 Then
            Return "Error WS07: La factura, el empresario y la fecha no coinciden. <b>Revise su ticket de compra</b>."
        End If

        If CDbl(Fila("Impreso")) <> 0 Then
            Return "Error WS08: La factura ha sido previamente impresa el dia: " & Format(CDate(Fila("FechaImpresion")), "dd/MM/yyyy HH:mm:ss")
        End If

        Dim sPath As String = ConfigurationManager.AppSettings("PathInvFactRep").ToString
        sPath &= "FacturaUnica1DigN01.rpt"
        Dim sFileName As String = sPrefijo & sNoPed & Format(Now, "_hhmmss") & ".pdf"
        Dim sRutaARchivo As String = ConfigurationManager.AppSettings("WhereToSavePDF").ToString & sFileName
        Dim miReporte As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim sRutaXml As String = ""

        Try
            If Not IO.File.Exists(sPath) Then
                Return "El archivo no existe en a ruta: " & sPath
            End If

            miReporte.Load(sPath, CrystalDecisions.Shared.OpenReportMethod.OpenReportByTempCopy)

            miReporte.SetParameterValue("@numfactura", sNoPed)
            miReporte.SetParameterValue("@prefijo", sPrefijo)
            miReporte.SetParameterValue("@cveempresario", sCveEmpresario)
            miReporte.SetParameterValue("@cvesucursal", CInt(Fila("SucFact")))

            miReporte.ExportToDisk(5, sRutaARchivo)

            miReporte.Close()
            miReporte.Dispose()

            If EnviarCorreo = "1" And Fila("estaTimbrado") = "0" Then 'Cualquier otro numero no envia correo
                'Enviar correo
                sRutaXml = Fila("DirXML").ToString

                Dim sFecha As String = Format(CDate(Fila("FA_FAC")), "dd/MMMM/yyyy")
                Dim sSubject As String = "Factura " & sPrefijo & sNoPed
                Dim sFileMaqueta As String = ConfigurationManager.AppSettings("WhereToRead").ToString & "MockMailFact.txt"
                Dim sBody As String = LeerArchivo1Time(sFileMaqueta)
                sBody = sBody.Replace("[FECHAFACTURA]", sFecha)
                sBody = sBody.Replace("[FECHAEXPORTAfull]", Now.ToString)
                sBody = sBody.Replace("[FACTNUM]", sPrefijo & sNoPed)

                If Fila("FormaVenta").ToString = "2" And Fila("Paqueteria").ToString = "E" Then
                    sBody = sBody.Replace("[LEYENDAPAQ]", "Puedes hacer el seguimiento de tu envío con tu número de rastreo largo <b>" & Fila("numguia").ToString & "</b> en la web de Estafeta: <br/>https://www.estafeta.com/Herramientas/Rastreo ")
                ElseIf Fila("FormaVenta").ToString = "2" And Fila("Paqueteria").ToString = "D" Then
                    sBody = sBody.Replace("[LEYENDAPAQ]", "Puedes hacer el seguimiento de tu envío con tu número de rastreo <b>" & Fila("numguia") & "</b> descargando la aplicación de Dostavista para dispositivos móviles: <br/><br/> *Android: <a href='https://play.google.com/store/apps/details?id=mx.delivery.client&referrer=af_tranid%3DtZ_qYl4vKjufE89BNb_-Ug%26shortlink%3D2bff276d%26pid%3Dwebsite_footer%26af_web_id%3D48f49ef5-41a2-4141-bf90-808e148efe1e-o'>Clic aquí</a><br/><br/> *iOS: <a href='https://apps.apple.com/MX/app/id1319174865?mt=8'>Clic aquí</a><br/>")
                ElseIf Fila("FormaVenta").ToString = "2" And Fila("Paqueteria").ToString = "H" Then
                    sBody = sBody.Replace("[LEYENDAPAQ]", "Puedes hacer el seguimiento de tu envío con tu número de rastreo <b>" & Fila("numguia") & "</b> en la web de DHL: <br/> http://www.dhl.com.mx/es/express/rastreo.html?AWB=" & Fila("numguia").ToString & "&brand=DHL")
                Else
                    sBody = sBody.Replace("[LEYENDAPAQ]", "")
                End If

                If Fila("FormaVenta").ToString <> "3" Then
                    sBody = sBody.Replace("[NUM_AUT]", ".")
                    sBody = sBody.Replace("[NUM_AUT_PRE]", "")
                    '[NUM_AUT_PRE]
                ElseIf Fila("FormaVenta").ToString = "3" Then
                    Dim sCadena As String
                    sCadena = "Su compra por internet en Forever Living Products México ha sido realizada <b>éxitosamente</b>."
                    sCadena &= "<br><b>Empresario</b>: " & Fila("CveEmpresario").ToString & " - " & Fila("Nombre").ToString
                    sCadena &= "<br><b>Total C.C.</b>: " & Fila("Puntos").ToString
                    sCadena &= "<br><b>Teléfono.</b>: " & Fila("Tel").ToString
                    sCadena &= "<br><b>Dirección envio:</b> " & Fila("dirEnvio").ToString
                    sCadena &= "<br><br>En breve recibirá un correo electrónico con el número de guía para que usted pueda realizar el rastreo de su pedido en el sitio <a target='_blank' href='http://www.estafeta.com.mx'>www.estafeta.com.mx</a> <br><br>"

                    sBody = sBody.Replace("[NUM_AUT]", " y con código de autorización <b>" & Fila("No_Aut") & "</b>.")
                    sBody = sBody.Replace("[NUM_AUT_PRE]", sCadena)
                End If


                '[FACTNUM]
                Dim sListaMails As String = Fila("Email").ToString
                Dim CCO As String = ""

                'Enviar correo a ventasFlpm cuando la compra se hace por 800 e Internet
                If Fila("CCO").ToString.Trim.Length > 0 And Fila("No_Aut").ToString.Trim.Length > 0 And (Fila("FormaVenta") = 3) Then
                    CCO = Fila("CCO").ToString
                End If

                sRES = "TODOOK"
                If Fila("Email").ToString.Trim.Length > 0 Then
                    sRES = EnviarMail(sSubject, sBody, sListaMails, CCO, sRutaARchivo & ";" & sRutaXml)
                    If sRES = "TODOOK" Then
                        'sRES = GuardarTicketImpreso(sCveEmpresario, sNoPed, sPrefijo, Fila("SucFact"), Fila("Email"))

                        Try
                            'Guardar bitacora
                            HsInfo = New Hashtable
                            HsInfo.Add("@idUsuario", 0)
                            HsInfo.Add("@CveSucursal", 0)
                            HsInfo.Add("@NotaVenta", 0)
                            HsInfo.Add("@Prefijo", sPrefijo)
                            HsInfo.Add("@NumFactura", sNoPed)
                            HsInfo.Add("@Observacion", "Impresion ticket|" & sCveEmpresario)
                            HsInfo.Add("@Origen", "")
                            HsInfo.Add("@IP", "")
                            HsInfo.Add("@PCName", "")
                            HsInfo.Add("@MacAddress", "")
                            HsInfo.Add("@OtrosDatos", sRES)

                            sRES = EjecutaSP("ws_spInsertaBitacora", HsInfo)
                        Catch ex As Exception

                        End Try

                    End If

                    Return sRES
                Else
                    Return "TODOOK" 'NO TIENE CORREO
                End If
            Else
                sRES = GuardarTicketImpreso(sCveEmpresario, sNoPed, sPrefijo, Fila("SucFact"), Fila("Email"))
                'devuelve URL
                Dim sURL As String = ConfigurationManager.AppSettings("LinkDownloadInvfact").ToString & sFileName
                Return sURL

            End If

        Catch ex As Exception
            Return "Error al exportar WS08: " & ex.Message.ToString & "<br>Rpt: " & sPath
        End Try

        Return "TODOOK"
    End Function

    Private Function GuardarTicketImpreso(ByVal sCveEmpresario As String,
                                          ByVal sNumFactura As String,
                                          ByVal sPrefijo As String,
                                          ByVal sCveSucursal As String,
                                          ByVal sEmail As String) As String

        Dim sRES As String
        Dim hsInfo As Hashtable
        hsInfo = New Hashtable
        hsInfo.Add("@CveEmpresario", CDbl(sCveEmpresario))
        hsInfo.Add("@NumFactura", CDbl(sNumFactura))
        hsInfo.Add("@prefijo", sPrefijo)
        hsInfo.Add("@cveSucursal", CInt(sCveSucursal))
        hsInfo.Add("@email", sEmail)

        sRES = EjecutaSP("spTck_InsImpreso", hsInfo)

        Return sRES

    End Function

    <WebMethod()>
    Public Function SetInfoFromClient(ByVal sXML As String) As String

        Dim ds As New DataSet
        Dim hsInfo As New Hashtable

        Try
            hsInfo.Add("@xml", sXML)
            ds = EjecutaDatasetSP("[sp400_InsXmlInfo]", hsInfo)
            If ds Is Nothing Then
                Return "Error WS05: El SP [sp400_InsXmlInfo] devolvió nothing."
            ElseIf ds.Tables(0).Rows.Count = 0 Then
                Return "Error WS06: El SP [sp400_InsXmlInfo] devolvió 0 registros."
            End If

            Dim Fila As DataRow
            Fila = ds.Tables(0).Rows(0)

            If Fila("Resp") = "TODOOK" Then
                Return "TODOOK"
            End If

            Return "TODOOK"
        Catch ex As Exception
            Return "Error 01A: " & ex.Message.ToString
        End Try

    End Function

    <WebMethod()>
    Public Function EnviarFactura(ByVal CveEmpresario As String,
                                ByVal Prefijo As String,
                                ByVal NumFactura As String,
                                ByVal CveSucursal As String,
                                ByVal EnviarCorreo As String) As String

        'CveEmpresario = "520000144080"
        'NumFactura = "45249"
        'CveSucursal = "1"
        'EnviarCorreo = "1"

        Dim hsInfo As Hashtable
        Dim ds As DataSet
        Dim sRES As String
        Dim sBody As String

        Try
            hsInfo = New Hashtable
            hsInfo.Add("@Prefijo", Prefijo)
            hsInfo.Add("@NumFactura", NumFactura)
            hsInfo.Add("@CveSucursal", CveSucursal)
            hsInfo.Add("@CveEmpresario", CveEmpresario)

            ds = EjecutaDatasetSP("spTck_GetFact", hsInfo)

            If ds Is Nothing Then
                Return "Error 002: Error en 'spTck_GetFact' devolvió nothing."
            End If

            If ds.Tables(0).Rows.Count = 0 Then
                Return "Error 002-A: No se encontró la factura. Revisar los datos."
            End If

            Dim Fila As Data.DataRow
            Fila = ds.Tables(0).Rows(0)

            Dim sHTML As String = ""
            Dim sEstilo As String
            sEstilo = ""
            sHTML &= "<div><img src='http://www.foreverecom.lat/images/emailheader.jpg' /></div>"
            sHTML &= "<div style='font-family: Tahoma,Century Gothic,sans-serif; font-size:10pt'>"
            sHTML &= "<br/>Estimado FBO <b>" & Fila("CveEmpresario") & " - " & Fila("Nombre") & "</b> se ha realizado su solicitud de compra <b>[FACTNUM]</b> del dia [FECHAFACTURA], " &
            "la cual se detalla a continuación: <br><br>"

            sHTML &= "<table style='background-color: #FFF;font-family: Tahoma; font-size: 10pt; color: #000;'>" &
            "<tr>" &
            "<td style='background-color: #1B0A2A;font-family: Tahoma; font-size: 10pt; color: #FFFFFF;'><b>Cantidad</b></td>" &
            "<td style='background-color: #1B0A2A;font-family: Tahoma; font-size: 10pt; color: #FFFFFF;'><b>Unidad</b></td>" &
            "<td style='background-color: #1B0A2A;font-family: Tahoma; font-size: 10pt; color: #FFFFFF;'><b>Código</b></td>" &
            "<td style='background-color: #1B0A2A;font-family: Tahoma; font-size: 10pt; color: #FFFFFF;'><b>Producto</b></td>" &
            "<td style='background-color: #1B0A2A;font-family: Tahoma; font-size: 10pt; color: #FFFFFF;'><b>Precio</b></td>" &
            "<td style='background-color: #1B0A2A;font-family: Tahoma; font-size: 10pt; color: #FFFFFF;'><b>Importe</b></td>" &
            "</tr>"
            For Each Fila2 As DataRow In ds.Tables(1).Rows
                sHTML &= "<tr>"
                sHTML &= "<td align='right'>" & Fila2("CantArt") & " </td>"
                sHTML &= "<td align='center'>" & Fila2("TipoUnidad") & " </td>"
                sHTML &= "<td align='right'>" & Format(Fila2("CveArt"), "000") & " </td>"
                sHTML &= "<td align='left'>" & Fila2("Descripcion") & " </td>"
                sHTML &= "<td align='right'>" & Format(CDbl(Fila2("PrecioArt")), "currency") & " </td>"
                sHTML &= "<td align='right'>" & Format(CDbl(Fila2("TotalArticulos")), "currency") & " </td>"
                sHTML &= "</tr>"
            Next
            sHTML &= "<tr><td colspan='6'><br><br><b>Tipo de Cambio:</b> $" & Fila("TipoCambio") & " <b>Total CC:</b> " & Fila("Puntos") & " <b>Total Pesos:</b> " & Format(CDbl(Fila("ImporteTotal")), "currency") & "</td></tr>"
            sHTML &= "</table>"

            sHTML &= "<table style='background-color: #FFF;font-family: Tahoma; font-size: 10pt; color: #000;'>"
            sHTML &= "<tr><td><br><b>Direccion de envio:</b></td></tr>"
            sHTML &= "<tr><td>" & Fila("EHADDR") & ", " & Fila("EHSHCT") & ", " & Fila("EHSHST") & " C.P. " & Fila("EHSHZP") & "</td></tr>"
            sHTML &= "<tr><td><br>Su número de guía es: <b>" & Fila("Numguia") & "</b></td></tr>"
            sHTML &= "<tr><td><br/>Para dar seguimiento al paquete en <b>ESTAFETA </b>  <a href='http://estafeta.azurewebsites.net/Tracking/searchByGet/?wayBillType=0&wayBill=" & Fila("Numguia") & "'>clic aqui</a></td></tr>"
            sHTML &= "<tr><td>Para ver el video de <b>guia </b> <a href='https://youtu.be/yYFYkzDjLVA'>clic aqui</a></td></tr>"
            sHTML &= "<tr><td>Además, para revisar la frecuencia de entrega en su localidad favor de <a href='http://www.estafeta.com/Frecuencia-de-entrega/'>consultar aqui</a></td></tr>"
            sHTML &= "<tr><td></td></tr>"
            sHTML &= "</table>"
            sHTML &= "</div>"


            Dim sPath As String = ConfigurationManager.AppSettings("PathInvFactRep").ToString
            sPath &= "FacturaUnica1DigN01.rpt"
            Dim sFileName As String = Fila("Prefijo") & Fila("NumFactura") & Format(Now, "_hhmmss") & ".pdf"
            Dim sRutaARchivo As String = ConfigurationManager.AppSettings("WhereToSavePDF").ToString & sFileName
            Dim miReporte As New CrystalDecisions.CrystalReports.Engine.ReportDocument

            If Not IO.File.Exists(sPath) Then
                Return "El archivo no existe en a ruta: " & sPath
            End If

            miReporte.Load(sPath, CrystalDecisions.Shared.OpenReportMethod.OpenReportByTempCopy)
            miReporte.SetParameterValue("@numfactura", Fila("NumFactura"))
            miReporte.SetParameterValue("@prefijo", Fila("Prefijo"))
            miReporte.SetParameterValue("@cveempresario", Fila("CveEmpresario"))
            miReporte.SetParameterValue("@cvesucursal", CInt(CveSucursal))
            miReporte.ExportToDisk(5, sRutaARchivo)
            miReporte.Close()
            miReporte.Dispose()

            If EnviarCorreo = "1" Then
                Dim sFecha As String = Format(CDate(Fila("FechaAlta")), "dd/MMMM/yyyy")
                Dim sSubject As String = "Factura " & Fila("Prefijo") & Fila("NumFactura")
                Dim sFileMaqueta As String = ConfigurationManager.AppSettings("WhereToRead").ToString & "MockMailFact.txt"
                sBody = LeerArchivo1Time(sFileMaqueta)
                sBody = sBody.Replace("[NUM_AUT_PRE]", sHTML)
                sBody = sBody.Replace("[FECHAFACTURA]", sFecha.ToUpper)
                sBody = sBody.Replace("[FECHAEXPORTAfull]", Now.ToString("dd/MMM/yyyy HH:mm").ToUpper)
                sBody = sBody.Replace("[FACTNUM]", Fila("Prefijo") & Fila("NumFactura"))
                If Fila("Paqueteria") = "E" Then
                    sBody = sBody.Replace("[LEYENDAPAQ]", "Puedes hacer el seguimiento de tu envío con tu número de rastreo <b>" & Fila("NumGuia") & "</b> en la web de Estafeta: <br/>https://www.estafeta.com/Herramientas/Rastreo ")
                ElseIf Fila("Paqueteria") = "H" Then
                    sBody = sBody.Replace("[LEYENDAPAQ]", "Puedes hacer el seguimiento de tu envío con tu número de rastreo <b>" & Fila("NumGuia") & "</b> en la web de DHL: <br/> http://www.dhl.com.mx/es/express/rastreo.html?AWB=" & Fila("NumGuia").ToString & "&brand=DHL")
                ElseIf Fila("Paqueteria") = "D" Then
                    sBody = sBody.Replace("[LEYENDAPAQ]", "Puedes hacer el seguimiento de tu envío con tu número de rastreo <b>" & Fila("NumGuia") & "</b> descargando la aplicación de Dostavista para dispositivos móviles: <br/><br/> *Android: <a href='https://play.google.com/store/apps/details?id=mx.delivery.client&referrer=af_tranid%3DtZ_qYl4vKjufE89BNb_-Ug%26shortlink%3D2bff276d%26pid%3Dwebsite_footer%26af_web_id%3D48f49ef5-41a2-4141-bf90-808e148efe1e-o'>Clic aquí</a><br/><br/> *iOS: <a href='https://apps.apple.com/MX/app/id1319174865?mt=8'>Clic aquí</a><br/>")
                Else
                    sBody = sBody.Replace("[LEYENDAPAQ]", "")
                End If
                'sBody &= "</font> </body>"
                sBody = sEstilo & sBody

                Dim sListaMails As String = Fila("Email").ToString
                Dim CCO As String = ""
                If (Fila("formaVenta") = "3" Or Fila("formaVenta") = "2") And Fila("CCO").ToString.Trim.Length > 0 Then
                    CCO = Fila("CCO")
                    CCO = CCO.ToString & ";arodriguez@foreverliving.com.mx"
                End If

                'sListaMails = "arodriguez@foreverliving.com.mx"
                'sRutaARchivo = ""
                'CCO = CCO.ToString & ";arodriguez@foreverliving.com.mx"

                sRES = ""
                If Fila("Email").ToString.Length > 0 Then
                    sRES = EnviarMail(sSubject, sBody, sListaMails, CCO, sRutaARchivo)
                End If

                If sRES <> "TODOOK" Then
                    Return sRES
                End If

            End If

            Return "TODOOK"
        Catch ex As Exception
            Return "Error WS01: " & ex.Message.ToString
        End Try
    End Function

    <WebMethod()>
    Public Function GetRpt(ByVal RptFullPath As String,
                            ByVal sXML_Param As String,
                            ByVal ExportFormat As String,
                            ByVal sMailToSend As String,
                            ByVal sMailSubject As String,
                            ByVal sMailBody As String,
                            ByVal PassPhrase As String) As String

        Dim sExtension As String
        Dim sRESP As String
        Dim sFileName As String
        Dim sNomReporte As String = System.IO.Path.GetFileNameWithoutExtension(RptFullPath)
        Dim sRutaArchivoSalida As String
        Dim miReporte As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim Col As String
        Dim hsInfo As New Hashtable

        Try
            Dim oDoc As New XmlDocument
            oDoc.LoadXml(sXML_Param)
            '//////////////////////
            For Each Nodo As System.Xml.XmlNode In oDoc.ChildNodes
                If Nodo.Name = "envelope" Then
                    For Each NodoHijo As System.Xml.XmlNode In Nodo.ChildNodes
                        If NodoHijo.Name = "param" Then
                            hsInfo.Add(NodoHijo.Attributes(0).Value, NodoHijo.Attributes(1).Value)
                        End If
                    Next
                End If
            Next
            '//////////////////////


            If PassPhrase <> "melquiades" Then
                Return "[Error 01]: PassPhrase no correcta."
            ElseIf Not IO.File.Exists(RptFullPath) Then
                Return "[Error 02]: Archivo .rpt no encontrado en la ruta especificada."
            End If

            If Not IsNumeric(ExportFormat) Then
                ExportFormat = "5"
            End If

            Select Case CInt(ExportFormat)
                Case 1
                    sExtension = ".rpt"
                Case 3
                    sExtension = ".doc"
                Case 4
                    sExtension = ".xls"
                Case 5
                    sExtension = ".pdf"
                Case Else
                    sExtension = ".dat"
            End Select

            'Preparar el nombre del archivo para la exportación
            If hsInfo.ContainsKey("returnThisFileName") Then
                sFileName = hsInfo("returnThisFileName") & sExtension
            Else
                sNomReporte = System.IO.Path.GetFileNameWithoutExtension(RptFullPath)
                sFileName = sNomReporte & "_ws" & Now.ToString("HHmmss") & sExtension
            End If

            'sRutaArchivoSalida = ConfigurationManager.AppSettings("WhereToSavePDF").ToString & sFileName
            If hsInfo.ContainsKey("DirectorioEscritura") Then
                sRutaArchivoSalida = hsInfo("DirectorioEscritura") & sFileName
            Else
                sRutaArchivoSalida = ConfigurationManager.AppSettings("WhereToSavePDF").ToString & sFileName
            End If


            'Si el archivo existe, eliminarlo.
            If IO.File.Exists(sRutaArchivoSalida) Then
                IO.File.Delete(sRutaArchivoSalida)
            End If

            'Cargar el reporte de crystal
            miReporte.Load(RptFullPath, CrystalDecisions.Shared.OpenReportMethod.OpenReportByTempCopy)

            'Agregar los parametros al reporte de crystal
            For iPos As Integer = 0 To miReporte.ParameterFields.Count - 1
                For Each KeyHash As Object In hsInfo.Keys
                    Col = KeyHash.ToString
                    If Col = miReporte.ParameterFields(iPos).Name Then
                        miReporte.SetParameterValue(Col, hsInfo(Col))
                        Exit For
                    End If
                Next
            Next


            miReporte.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.DefaultPaperSize

            'Exportar el crystal al forma especificado en la ruta especificado.
            miReporte.ExportToDisk(CInt(ExportFormat), sRutaArchivoSalida)

            'Liberar los recursos del crystal.
            miReporte.Close()
            miReporte.Dispose()

            'Revisar si responde una URL o envia un mail
            If sMailToSend = "URL" Then
                'Devoler URL
                sRESP = ConfigurationManager.AppSettings("LinkDownloadInvfact").ToString & sFileName
            Else
                'intentar mandar un mail con el archivo atado.
                sRESP = EnviarMail(sMailSubject, sMailBody, sMailToSend, "", sRutaArchivoSalida)
            End If

            If hsInfo.ContainsKey("DirectorioEscritura") Then
                sRESP = sRutaArchivoSalida
            End If

            Return sRESP
        Catch ex As Exception
            miReporte.Close()
            miReporte.Dispose()
            Return "[Error -99]: " & ex.Message.ToString
        End Try
    End Function

    <WebMethod()>
    Public Function getRFC(ByVal strNombre As String,
                            ByVal strPaterno As String,
                            ByVal strMaterno As String,
                            ByVal dteFechaNacimiento As String) As String

        'Esta rutina genera el RFC. Datos de entrada:

        'strNombre: Tipo String Nombre de pila Dato valido requerido.
        'strPaterno: Tipo String Apellido paterno Por lo menos un
        'strMaterno: Tipo String Apellido materno apellido es requerido.
        'dteFechaNacimiento: Tipo Date


        Dim strFecha As String
        Dim strRFC As String = ""
        Dim strNombreOriginal As String
        Dim strPaternoOriginal As String
        Dim strMaternoOriginal As String

        Try

            'Consigue la fecha.
            strFecha = Format(CDate(dteFechaNacimiento), "yyMMdd")

            'El RFC se calcula a base de letras vocales
            'sin acento, elimina cualquier acento dentro de las variables
            strNombre = RFCFiltraAcentos(strNombre)
            strPaterno = RFCFiltraAcentos(strPaterno)
            strMaterno = RFCFiltraAcentos(strMaterno)

            'Guarda el nombre original para calcular
            'la homoclave.
            strNombreOriginal = strNombre
            strPaternoOriginal = strPaterno
            strMaternoOriginal = strMaterno

            'Elimina palabras sobrantes dentro de los nombres.
            RFCFiltraNombres(strNombre, strPaterno, strMaterno)

            'Ahora el siguiente paso es determinar como se va a
            'calcular el RFC. Existen reglas:
            '
            ' 1. Se dan los tres nombres.
            ' 2. Se da solo un nombre.
            ' 3. El apellido dado solo tiene 3 o menos letras.
            Select Case True
                Case Len(strPaterno) > 0 And Len(strMaterno) > 0
                    'Los tres nombres existen, procede.
                    'Determina si el apellido paterno tiene
                    'menos de 3 letras.
                    If Len(strPaterno) < 3 Then
                        'Calcula el RFC tomando en cuenta un apellido corto.
                        strRFC = RFCApellidoCorto(strNombre, strPaterno, strMaterno, strFecha)
                    Else
                        'Calcula el RFC.
                        strRFC = RFCArmalo(strNombre, strPaterno, strMaterno, strFecha)
                    End If

                Case Len(strPaterno) = 0 Or Len(strMaterno) = 0
                    'Uno de ellos esta vacio.
                    strRFC = RFCUnApellido(strNombre, strPaterno, strMaterno, strFecha)

            End Select

            'El RFC tentativo ya esta armado. Ahora elimina
            'cualquier palabra posiblemente ofensiva.
            strRFC = RFCQuitaProhibidas(strRFC)

            'Ya tengo el RFC, ahora solo falta la homoclave y el
            'digito verificador.
            strRFC = strRFC & RFCHomoclave(strNombreOriginal, strPaternoOriginal, strMaternoOriginal)

            'Por ultimo, calcula el digito verificador.
            Return strRFC & RFCDigitoVerificador(strRFC)
        Catch ex As Exception
            Return "Error[01a]: " & ex.Message.ToString
        End Try
    End Function

    <WebMethod()>
    Public Function cfdiIngreso_validar(ByVal IdBono As String,
                                        ByVal sDirBase As String) As String

        Dim hsInfo As Hashtable
        Dim ds As DataSet
        Dim sRES As String = ""
        Dim sRutaXML As String = ""
        Dim idPago As Double = 0
        Try
            'Escribir_Archivo("IdBono: " & IdBono, "D:\ftproot\DHK\bita.txt")
            'Escribir_Archivo("sDirBase: " & sDirBase, "D:\ftproot\DHK\bita.txt")

            EscribirBitacora("-------------------------------------------------------------------------", "C:\inetpub\wwwroot\wsInterfaz164\Out\bitacora\bitacora.txt")
            EscribirBitacora("IdBono: " & IdBono & vbCrLf & "sDirBase: " & sDirBase, "C:\inetpub\wwwroot\wsInterfaz164\Out\bitacora\bitacora.txt")

            'Obtener los valores del bono desde la BDD
            hsInfo = New Hashtable
            hsInfo.Add("@idBono", CDbl(IdBono))
            ds = FillDataSet("spSel_BonoId", sRES, hsInfo)

            If Not sRES = "TODOOK" Then
                Return "Error[02]: " & sRES
            End If

            If ds Is Nothing Then
                Return "Error[03]: Procedimiento fallido."
            End If

            If ds.Tables(0).Rows.Count = 0 Then
                Return "Error[04]: No se encuentra el identficador del bono."
            End If

            Dim drBONO As DataRow = ds.Tables(0).Rows(0)

            idPago = drBONO("idPago")

            sRutaXML = sDirBase & ds.Tables(0).Rows(0)("rutaArchivo")
            sRutaXML = sRutaXML.Replace("\\", "\")

            If Not IO.File.Exists(sRutaXML) And sDirBase <> "CFDI-TABLA" Then
                Return "Error[01]: El archivo no existe en la ruta especificada."
            End If

            If sDirBase = "CFDI-TABLA" Then
                Dim rutaRaiz As String = "" & drBONO("idAnio") & "\" & Format(CInt(drBONO("idMes")), "0")
                Dim sFileName As String = drBONO("cveEmpresario") & "_" & MonthName(CInt(drBONO("idMes")), True).ToUpper & drBONO("idAnio") & ".xml"
                Dim NewDirectorio As String = ConfigurationManager.AppSettings("dirBaseXmlCfdi").ToString & "" & rutaRaiz

                If Not IO.Directory.Exists(NewDirectorio) Then
                    IO.Directory.CreateDirectory(NewDirectorio)
                End If

                sRutaXML = ConfigurationManager.AppSettings("dirBaseXmlCfdi").ToString & "" & rutaRaiz & "\" & sFileName
                'dirBaseXmlCfdi C:\Inetpub\wwwroot\SicobonV2\Uploads\cfdi\
                Dim sResCrearArch As String
                sResCrearArch = Escribir_Archivo(drBONO("xmlCfdiString"), sRutaXML)

                If Not IO.File.Exists(sRutaXML) Then
                    EscribirBitacora("El archivo no se pudo guardar: " & sResCrearArch & vbCrLf & "Ruta: " & sRutaXML, "C:\inetpub\wwwroot\wsInterfaz164\Out\bitacora\bitacora.txt")
                Else
                    EscribirBitacora("Archivo generado: " & sRutaXML, "C:\inetpub\wwwroot\wsInterfaz164\Out\bitacora\bitacora.txt")
                End If

                ' Return sRutaXML
            End If


            'sRutaXML = "C:\Inetpub\wwwroot\SicobonV2\Uploads\cfdi\2014\2\mar14000031751.xml"
            'Leer todo el contenido del archivo para validar el contenido, por si trae un caracter raro
            Dim sContenido As String = LeerArchivo1Time(sRutaXML)

            If sContenido.StartsWith("?<?xml") Then
                sContenido = sContenido.Replace("?<?xml", "<?xml")
                Escribir_Archivo(sContenido, sRutaXML) 'Sobre-escribir el archivo con el nuevo contenido
            End If

            'sRutaXML = "D:\ftproot\dhk\mar14000430.xml"

            'Obtener en un data table toda la información del bono y en otro la información del XML
            hsInfo = New Hashtable
            hsInfo.Add("@idUsuario", -1)
            hsInfo.Add("@idBono", CDbl(IdBono))
            hsInfo.Add("@rutaXml", sRutaXML)
            ds = FillDataSet("[spSel_CfdiValidar]", sRES, hsInfo)

            If Not sRES = "TODOOK" Then
                Return "Error[VAL02]: " & sRES
            End If

            If ds Is Nothing Then
                Return "Error[VAL03]: Procedimiento fallido."
            End If

            If ds.Tables(0).Rows.Count = 0 Then
                Return "Error[VAL04]: Problema con la importación de XML-CFDI."
            End If

            If ds.Tables(1).Rows.Count = 0 Then
                Return "Error[VAL05]: Problema con la lectura del bono."
            End If


            'saveDStoFile(ds, "D:\ftproot\DHK\dato_" & IdBono & ".xls")

            Dim drCFDI As DataRow = ds.Tables(0).Rows(0)
            Dim drBONOVAL As DataRow = ds.Tables(1).Rows(0)
            Dim drFiscales As DataRow = ds.Tables(2).Rows(0)

            If Not drCFDI("tipoDeComprobante").ToUpper.Equals("INGRESO") And Not drCFDI("tipoDeComprobante").ToUpper.Equals("I") Then
                Return "El comprobante debe ser de tipo INGRESO=" & drCFDI("tipoDeComprobante").ToUpper
            ElseIf Not drCFDI("EmisorRFC").Trim = drBONOVAL("rfc").ToString.Trim Then
                Return "El RFC del emisor del comprobante [" & drCFDI("EmisorRFC").Trim & "] no coincide con el registrado en FLP México [" & drBONOVAL("rfc").ToString.Trim & "]"
            End If

            'Revisar los datos del receptor (FLP MEX)
            'If Not quitarNoValidos(drCFDI("ReceptorNombre").ToString.Trim).StartsWith(drFiscales("razSoc").ToString.Trim) Then
            '    Return "La razón social del receptor es incorrecta. El valor correcto es: " & drFiscales("razSoc").ToString & ". Valor incorrecto: " & drCFDI("ReceptorNombre").ToString.Trim
            'End If

            If Not drCFDI("version4_0").ToString.Trim = "4.0" Then
                Return "La versión del CFDI debe ser 4.0. Favor de corregir."
            End If

            If Not drCFDI("ReceptorRFC").ToString.Trim = drFiscales("rfc").ToString.Trim Then
                Return "RFC del receptor es incorrecta. El valor correcto es: " & drFiscales("RFC").ToString & ". Valor incorrecto: " & drCFDI("ReceptorRFC").ToString.Trim
            End If


            'Realizar los importes del bono
            If Not CDbl(drCFDI("importe")) = CDbl(drBONOVAL("importe")) Then
                If Not CDbl(drCFDI("subTotal")) = CDbl(drBONOVAL("importe")) Then
                    Return "El importe|valor unitario|subtotal del CFDI es incorrecto. El valor correcto es: " & CDbl(drBONOVAL("importe")) '& ". Valor incorrecto: " & drCFDI("Importe").ToString
                End If
            ElseIf CDate(drCFDI("FechaTimbrado")) <= CDate(drBONOVAL("FechaAlta")) Then
                Return "La fecha de timbrado del CFDI es menor o igual a la fecha de emisión del bono."
            End If

            If drCFDI("total") > Math.Abs(CDbl(drBONOVAL("total")) + CDbl(7)) Or drCFDI("total") < Math.Abs(CDbl(drBONOVAL("total")) - CDbl(7)) Then
                Return "El total de la factura es incorrecto, sobrepasa el margen de aceptación."
            End If

            Dim diff As Double

            diff = Math.Abs(CDbl(drBONOVAL("IVA")) - CDbl(drCFDI("IVA")))
            'If diff > 2 Then
            '    Return "El monto del IVA es incorrecto. El valor correcto es: " & drBONOVAL("IVA").ToString
            'End If

            diff = Math.Abs(CDbl(drBONOVAL("retencion")) - CDbl(drCFDI("ISR")))
            'If diff > 2 Then
            '    Return "El monto del ISR es incorrecto. El valor correcto es: " & drBONOVAL("retencion").ToString
            'End If

            diff = Math.Abs(CDbl(drBONOVAL("total")) - CDbl(drCFDI("total")))
            If diff > 7 Then
                Return "El total a pagar es incorrecto. El valor correcto es: " & drBONOVAL("total").ToString
            End If

            Dim sSQL_tributarios As String = "update SicobonV2.dbo.pagos set obligfiscales = 1 where idpago = " & idPago.ToString & ";"
            EscribirBitacora("sSQL_tributarios: " & sSQL_tributarios, "C:\inetpub\wwwroot\wsInterfaz164\Out\bitacora\bitacora.txt")
            EjecutaDatasetQry(sSQL_tributarios)


            Return "TODOOK"
        Catch ex As Exception
            EscribirBitacora("Error [A01]: " & ex.Message.ToString, "C:\inetpub\wwwroot\wsInterfaz164\Out\bitacora\bitacora.txt")
            Return "Error [A01]: " & ex.Message.ToString
        End Try
    End Function

    <WebMethod()>
    Public Function qrCode_get(ByVal sInput As String,
                               ByVal ReturnFileName As String,
                               ByVal UrlOrDireccion As String,
                               ByVal contrasena As String) As String


        Try
            Dim qrEncoder As New Gma.QrCodeNet.Encoding.QrEncoder(Gma.QrCodeNet.Encoding.ErrorCorrectionLevel.H)
            Dim qrCode As New Gma.QrCodeNet.Encoding.QrCode
            qrEncoder.TryEncode(sInput, qrCode)

            Dim wRenderer As New WriteableBitmapRenderer(New FixedModuleSize(2, QuietZoneModules.Two))
            Dim ms As New IO.MemoryStream
            wRenderer.WriteToStream(qrCode.Matrix, ImageFormatEnum.JPEG, ms)

            Dim sFileName As String = ReturnFileName

            If sFileName.Trim.Length = 0 Then
                sFileName = Now.ToString("HHmmss") & ".jpg"
            Else
                sFileName = sFileName & ".jpg"
            End If

            Dim sWhereToSave As String = ConfigurationManager.AppSettings("WhereToSavePDF").ToString & sFileName

            Dim myImagen As Drawing.Image
            myImagen = Drawing.Image.FromStream(ms)
            myImagen.Save(sWhereToSave)

            If UrlOrDireccion.ToUpper = "URL" Then
                Return ConfigurationManager.AppSettings("LinkDownloadInvfact").ToString & sFileName
            Else
                Return sWhereToSave
            End If

        Catch ex As Exception
            Return "ERROR: " & ex.Message.ToString
        End Try

    End Function

End Class