Imports System.Xml

Public Class factura
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oCompany As SAPbobsCOM.Company
    Private oBusinessForm As SAPbouiCOM.Form
    Public oFilters As SAPbouiCOM.EventFilters
    Public oFilter As SAPbouiCOM.EventFilter
    Private oMatrix As SAPbouiCOM.Matrix        ' Global variable to handle matrixes

    Private AddStarted As Boolean                ' Flag that indicates "Add" process started

    Private RedFlag As Boolean                   ' RedFlag when true indicates an error during "Add" process


#Region "Single Sign On"

    Private Sub SetApplication()


        AddStarted = False

        RedFlag = False

        '*******************************************************************

        '// Use an SboGuiApi object to establish connection

        '// with the SAP Business One application and return an

        '// initialized application object

        '*******************************************************************
        Try
            Dim SboGuiApi As SAPbouiCOM.SboGuiApi

            Dim sConnectionString As String

            SboGuiApi = New SAPbouiCOM.SboGuiApi

            '// by following the steps specified above, the following

            '// statement should be sufficient for either development or run mode
            If Environment.GetCommandLineArgs.Length > 1 Then
                sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
            Else
                sConnectionString = Environment.GetCommandLineArgs.GetValue(0)
            End If

            'sConnectionString = Environment.GetCommandLineArgs.GetValue(1) '"0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"

            '// connect to a running SBO Application

            SboGuiApi.Connect(sConnectionString)

            '// get an initialized application object

            SBO_Application = SboGuiApi.GetApplication()
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString)
        End Try


    End Sub

    Private Function SetConnectionContext() As Integer

        Dim sCookie As String

        Dim sConnectionContext As String

        Dim lRetCode As Integer

        Try

            '// First initialize the Company object

            oCompany = New SAPbobsCOM.Company

            '// Acquire the connection context cookie from the DI API.

            sCookie = oCompany.GetContextCookie

            '// Retrieve the connection context string from the UI API using the

            '// acquired cookie.

            sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie)

            '// before setting the SBO Login Context make sure the company is not

            '// connected

            If oCompany.Connected = True Then

                oCompany.Disconnect()

            End If

            '// Set the connection context information to the DI API.

            SetConnectionContext = oCompany.SetSboLoginContext(sConnectionContext)

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

    End Function

    Private Function ConnectToCompany() As Integer

        '// Establish the connection to the company database.

        ConnectToCompany = oCompany.Connect

    End Function


    Private Sub Class_Init()
        Try
            'Dim wsCodifid As New WSCofidi.Service1
            'Using fileReader As New FileIO.TextFieldParser(Application.StartupPath & "\(NC) No.35.xml")
            'Dim respuesta = wsCodifid.GeneraDTE("00000REM01", "", "REMISA", "12345678", fileReader.ReadToEnd.ToString, "02", "")
            'Dim xmlResp As New XmlDocument
            'xmlResp.LoadXml(respuesta)
            ' Dim respNode = xmlResp.DocumentElement.SelectSingleNode("/EI/INVOICES/INVOICE")
            'Dim oFolioFiscal = respNode.Attributes("folio_fiscal")
            'Dim oSerieFiscal = respNode.Attributes("serie_fiscal")
            'Dim oSello = respNode.Attributes("sello")
            ' Dim errorMess = respNode.Attributes("error_message")
            ' xmlResp.Save(Application.StartupPath & "\Respuesta(NC) No.35.xml")

            'End Using
            'UDT_UF.Company = Me.oCompany
            'cargarInicial(oCompany, SBO_Application)
            '//*************************************************************

            '// set SBO_Application with an initialized application object

            '//*************************************************************
            Dim wsCodifi As New WSCofidi.Service1
            Using fileReader As New FileIO.TextFieldParser(Application.StartupPath & "\(F) No." & "1484" & ".xml")
                Dim respuesta = wsCodifi.GeneraDTE("0000000059", "", "admin.remisa", "R3m1s@!", fileReader.ReadToEnd.ToString, "02", "")
                Dim xmlResp As New XmlDocument
                xmlResp.LoadXml(respuesta)
                Dim respNode = xmlResp.DocumentElement.SelectSingleNode("/EI/INVOICES/INVOICE")
                Dim oFolioFiscal = respNode.Attributes("folio_fiscal")
                Dim oSerieFiscal = respNode.Attributes("serie_fiscal")
                Dim oSello = respNode.Attributes("sello")
                Dim errorMess = respNode.Attributes("error_message")
                Dim oResolu = respNode.Attributes("noAprobacion")
                xmlResp.Save(Application.StartupPath & "\Respuesta(F)" & "1484" & ".xml")
                SetApplication()
            End Using

            '//*************************************************************

            '// Set The Connection Context

            '//*************************************************************

            If Not SetConnectionContext() = 0 Then

                SBO_Application.MessageBox("Failed setting a connection to DI API")

                End ' Terminating the Add-On Application

            End If

            '//*************************************************************

            '// Connect To The Company Data Base

            '//*************************************************************

            If Not ConnectToCompany() = 0 Then

                SBO_Application.MessageBox("Failed connecting to the company's Data Base")

                End ' Terminating the Add-On Application

            End If

            '//*************************************************************

            '// send an "hello world" message

            '//*************************************************************

            SBO_Application.SetStatusBarMessage("DI Connected To: " & oCompany.CompanyName & vbNewLine & "Add-on is loaded", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            SetNewItems()

            'SetNewTax("01", "512 0% a 22 % pago al exterior", SAPbobsCOM.WithholdingTaxCodeCategoryEnum.wtcc_Invoice, SAPbobsCOM.WithholdingTaxCodeBaseTypeEnum.wtcbt_Net, 100, "512", "1-1-010-10-000")
            'SetNewTax("02", "513 0% a 22 % pago al exterior", SAPbobsCOM.WithholdingTaxCodeCategoryEnum.wtcc_Invoice, SAPbobsCOM.WithholdingTaxCodeBaseTypeEnum.wtcbt_Net, 100, "513", "1-1-010-10-000")
            'SetNewTax("03", "513A 0% a 22 % pago al exterior", SAPbobsCOM.WithholdingTaxCodeCategoryEnum.wtcc_Invoice, SAPbobsCOM.WithholdingTaxCodeBaseTypeEnum.wtcbt_Net, 100, "513A", "_SYS00000000128")
            'SetNewTax("04", "514 0% a 22 % pago al exterior", SAPbobsCOM.WithholdingTaxCodeCategoryEnum.wtcc_Invoice, SAPbobsCOM.WithholdingTaxCodeBaseTypeEnum.wtcbt_Net, 100, "514", "_SYS00000000128")

            TOOLS.SBOApplication = Me.SBO_Application
            Dim doc As New XmlDocument
            ' doc.LoadXml(pVal.ObjectKey)
            'Dim docEntrynode = doc.DocumentElement.SelectSingleNode("/DocumentParams/DocEntry")
            'generarXML("1158", "13")


            ' Dim wsCodifi As New WSCofidi.Service1
            'Dim oDocEnviar As New XmlDocument
            ' oDocEnviar.Load(Application.StartupPath & "\(F) No.18.xml")
            'Dim respuesta = wsCodifi.GeneraDTE("00000REM01", "", "REMISA", "12345678", oDocEnviar.InnerText, "02", "")
            'UDT_UF.Company = Me.oCompany
            'cargarInicial(oCompany, SBO_Application)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Sub

#End Region



    Public Sub New()

        MyBase.New()

        Class_Init()

        'AddMenuItems()

        'SetFilters()


    End Sub

    Private Sub SetFilters()

        '// Create a new EventFilters object

        oFilters = New SAPbouiCOM.EventFilters



        '// add an event type to the container

        '// this method returns an EventFilter object

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)
        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK)
        'oFilter = oFilter.Add(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK)
        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)

        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_VALIDATE)

        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS)

        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN)

        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)

        ' oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK)
        ' oFilter.AddEx("60006") 'Quotation Form



        '// assign the form type on which the event would be processed

        ' oFilter.AddEx("134") 'Quotation Form
        'oFilter.AddEx("141")
        'oFilter.AddEx("-141")
        oFilter.AddEx("133")
        oFilter.AddEx("179")
        'oFilter.AddEx("60004")
        'oFilter.AddEx("-133")
        'oFilter.AddEx("-181")
        'oFilter.AddEx("181")
        'oFilter.AddEx("-65303")
        'oFilter.AddEx("65303")
        'oFilter.AddEx("65306")
        'oFilter.AddEx("-65306")

        'oFilter.AddEx("-179")
        'oFilter.AddEx("139") 'Orders Form
        'oFilter.AddEx("133") 'Invoice Form
        'oFilter.AddEx("169") 'Main Menu
        SBO_Application.SetFilter(oFilters)

    End Sub

    Private Sub SBO_Application_DATAEVENT(ByRef pVal As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.FormDataEvent
        Dim docentry As String = ""
        Dim type As String = ""
        Try
            If pVal.FormTypeEx = "141" And pVal.BeforeAction = False And pVal.ActionSuccess = True Then
            End If
            If pVal.FormTypeEx = "133" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                'MessageBox.Show(pVal.ObjectKey)
                Dim doc As New XmlDocument
                doc.LoadXml(pVal.ObjectKey)
                Dim docEntrynode = doc.DocumentElement.SelectSingleNode("/DocumentParams/DocEntry")
                docEntry = docEntrynode.InnerText
                Type = "13"
                generarXML(docEntrynode.InnerText, "13")
                Dim wsCodifi As New WSCofidi.Service1
                Using fileReader As New FileIO.TextFieldParser(Application.StartupPath & "\(F) No." & docEntrynode.InnerText & ".xml")
                    Dim respuesta = wsCodifi.GeneraDTE("0000000059", "", "admin.remisa", "R3m1s@!", fileReader.ReadToEnd.ToString, "02", "")
                    Dim xmlResp As New XmlDocument
                    xmlResp.LoadXml(respuesta)
                    Dim respNode = xmlResp.DocumentElement.SelectSingleNode("/EI/INVOICES/INVOICE")
                    Dim oFolioFiscal = respNode.Attributes("folio_fiscal")
                    Dim oSerieFiscal = respNode.Attributes("serie_fiscal")
                    Dim oSello = respNode.Attributes("sello")
                    Dim errorMess = respNode.Attributes("error_message")
                    Dim oResolu = respNode.Attributes("noAprobacion")
                    xmlResp.Save(Application.StartupPath & "\Respuesta(F)" & docEntrynode.InnerText & ".xml")
                    Dim oRecord As SAPbobsCOM.Recordset
                    oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim errorenvio As String
                    If oSello.Value = "" Then
                        errorenvio = "No se pudo entregar Factura revisar XML de respuesta"

                    Else
                        errorenvio = ""
                    End If
                    oRecord.DoQuery("CALL FAC_UPDATE(' F=" & oFolioFiscal.Value & "- S=" & oSerieFiscal.Value & "- R=" & oResolu.Value & "','" & oSello.Value & "','" & docEntrynode.InnerText & "','" & errorenvio & "')")
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
                    oRecord = Nothing
                    GC.Collect()
                End Using

            End If
            If pVal.FormTypeEx = "179" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                Dim doc As New XmlDocument
                doc.LoadXml(pVal.ObjectKey)
                Dim docEntrynode = doc.DocumentElement.SelectSingleNode("/DocumentParams/DocEntry")
                docEntry = docEntrynode.InnerText
                Type = "14"
                generarXML(docEntrynode.InnerText, "14")
                Dim wsCodifi As New WSCofidi.Service1
                Using fileReader As New FileIO.TextFieldParser(Application.StartupPath & "\(NC) No." & docEntrynode.InnerText & ".xml")
                    Dim respuesta = wsCodifi.GeneraDTE("0000000059", "", "admin.remisa", "R3m1s@!", fileReader.ReadToEnd.ToString, "02", "")
                    Dim xmlResp As New XmlDocument
                    xmlResp.LoadXml(respuesta)
                    Dim respNode = xmlResp.DocumentElement.SelectSingleNode("/EI/INVOICES/INVOICE")
                    Dim oFolioFiscal = respNode.Attributes("folio_fiscal")
                    Dim oSerieFiscal = respNode.Attributes("serie_fiscal")
                    Dim oSello = respNode.Attributes("sello")
                    Dim errorMess = respNode.Attributes("error_message")
                    Dim oResolu = respNode.Attributes("noAprobacion")
                    xmlResp.Save(Application.StartupPath & "\Respuesta(NC)" & docEntrynode.InnerText & ".xml")
                    Dim oRecord As SAPbobsCOM.Recordset
                    oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim errorenvio As String
                    If oSello.Value = "" Then
                        errorenvio = "No se pudo entregar Factura revisar XML de respuesta"

                    Else
                        errorenvio = ""
                    End If
                    oRecord.DoQuery("CALL FAC_UPDATE_ORIN(' F=" & oFolioFiscal.Value & "- S=" & oSerieFiscal.Value & "- R=" & oResolu.Value & "','" & oSello.Value & "','" & docEntrynode.InnerText & "','" & errorenvio & "')")
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
                    oRecord = Nothing
                    GC.Collect()
                End Using
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Dim Sql As String = ""
            Dim oRecord As SAPbobsCOM.Recordset
            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If type = "14" Then
                Sql = "CALL FAC_UPDATE_ORIN('','','" & docentry & "','" & ex.Message & "')"
            Else
                Sql = "CALL FAC_UPDATE('','','" & docentry & "','" & ex.Message & "')"
            End If
            oRecord.DoQuery(Sql)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecord)
            oRecord = Nothing
            GC.Collect()
        End Try


    End Sub

    Private Sub generarXML(DocEntry As String, p2 As String)
        Try
            Dim oRecordEnc As SAPbobsCOM.Recordset
            Dim doc As New XmlDocument
            oRecordEnc = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim SQL = "CALL FAC_ENCABEZADO_FACE('" & p2 & "'," & Integer.Parse(DocEntry) & ")"
            oRecordEnc.DoQuery(SQL)
            Dim encabezado As String = ""
            If p2 = "13" Then
                encabezado = "(F) No."
            Else
                encabezado = "(NC) No."
            End If
            Dim writer As New XmlTextWriter(encabezado & DocEntry & ".xml", System.Text.Encoding.UTF8)
            writer.WriteStartDocument(True)
            writer.Formatting = Formatting.Indented
            writer.Indentation = 2
            writer.WriteStartElement("EI")
            writer.WriteStartElement("INVOICES")
            writer.WriteStartElement("INVOICE")

            If oRecordEnc.RecordCount > 0 Then
                While oRecordEnc.EoF = False
                    writer.WriteStartElement("HEADER")
                    writer.WriteStartElement("E01")
                    writer.WriteAttributeString("FolioInterno", oRecordEnc.Fields.Item(1).Value)
                    writer.WriteAttributeString("Fecha", Date.Now.ToString("yyyy-MM-dd 00:00:00"))
                    writer.WriteAttributeString("SubTotal", oRecordEnc.Fields.Item(3).Value)
                    writer.WriteAttributeString("Descuento", oRecordEnc.Fields.Item(4).Value)
                    writer.WriteAttributeString("Total", oRecordEnc.Fields.Item(5).Value)
                    writer.WriteAttributeString("EstadoDocumento", oRecordEnc.Fields.Item(6).Value)
                    writer.WriteAttributeString("TipoDeComprobante", oRecordEnc.Fields.Item(7).Value)
                    writer.WriteEndElement()
                    writer.WriteStartElement("E02")
                    writer.WriteAttributeString("Nit", oRecordEnc.Fields.Item(8).Value)
                    writer.WriteAttributeString("NoCliente", oRecordEnc.Fields.Item(9).Value)
                    writer.WriteAttributeString("Nombre", oRecordEnc.Fields.Item(10).Value)
                    writer.WriteEndElement()
                    writer.WriteStartElement("E03")
                    writer.WriteAttributeString("Localidad", oRecordEnc.Fields.Item(11).Value)
                    writer.WriteAttributeString("Pais", oRecordEnc.Fields.Item(12).Value)
                    writer.WriteEndElement()
                    writer.WriteStartElement("E04")
                    writer.WriteAttributeString("Impuesto", oRecordEnc.Fields.Item(13).Value)
                    writer.WriteEndElement()
                    writer.WriteStartElement("EA1")
                    writer.WriteAttributeString("CustomField15", oRecordEnc.Fields.Item(14).Value)
                    writer.WriteAttributeString("Moneda", oRecordEnc.Fields.Item(15).Value)
                    writer.WriteAttributeString("TipoCambio", oRecordEnc.Fields.Item(16).Value)
                    writer.WriteEndElement()
                    writer.WriteStartElement("C03")
                    writer.WriteAttributeString("Dato3", "")
                    writer.WriteEndElement()
                    writer.WriteStartElement("C04")
                    writer.WriteAttributeString("Dato4", "")
                    writer.WriteEndElement()
                    writer.WriteStartElement("C05")
                    writer.WriteAttributeString("Dato5", "")
                    writer.WriteEndElement()
                    writer.WriteStartElement("C06")
                    writer.WriteAttributeString("Dato6", "")
                    writer.WriteEndElement()
                    oRecordEnc.MoveNext()
                    'En HEADER
                    writer.WriteEndElement()
                End While
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordEnc)
            oRecordEnc = Nothing
            GC.Collect()

            Dim orecordDet As SAPbobsCOM.Recordset
            orecordDet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            orecordDet.DoQuery("CALL FAC_DETALLE_FACE('" & p2 & "'," & Integer.Parse(DocEntry) & ")")
            If orecordDet.RecordCount > 0 Then
                writer.WriteStartElement("DETAILS")
                While orecordDet.EoF = False
                    writer.WriteStartElement("D01")
                    writer.WriteAttributeString("Cantidad", orecordDet.Fields.Item(0).Value)
                    writer.WriteAttributeString("Descripcion", orecordDet.Fields.Item(1).Value)
                    writer.WriteAttributeString("Importe", orecordDet.Fields.Item(2).Value)
                    writer.WriteAttributeString("NoIdentificacion", orecordDet.Fields.Item(3).Value)
                    writer.WriteAttributeString("Unidad", orecordDet.Fields.Item(4).Value)
                    writer.WriteAttributeString("ValorUnitario", orecordDet.Fields.Item(5).Value)
                    writer.WriteStartElement("DA6")
                    writer.WriteAttributeString("Importe", orecordDet.Fields.Item(6).Value)
                    writer.WriteAttributeString("Impuesto", orecordDet.Fields.Item(7).Value)
                    writer.WriteAttributeString("Tasa", orecordDet.Fields.Item(8).Value)
                    writer.WriteEndElement()
                    writer.WriteStartElement("DA8")
                    writer.WriteAttributeString("Porcentaje", Math.Abs(Double.Parse(orecordDet.Fields.Item(9).Value)))
                    writer.WriteAttributeString("Descripcion", orecordDet.Fields.Item(10).Value)
                    writer.WriteAttributeString("Importe", Math.Abs(Double.Parse(orecordDet.Fields.Item(11).Value)))
                    writer.WriteEndElement()

                    'Fin D01
                    writer.WriteEndElement()
                    orecordDet.MoveNext()
                End While
                'Fin Details
                writer.WriteEndElement()
            End If



            'Fin INVOICE
            writer.WriteEndElement()
            'Fin INVOICES
            writer.WriteEndElement()
            'Fin EI
            writer.WriteEndElement()
            writer.Close()
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, True)
        End Try
    End Sub
    Private Sub createNode(ByVal pID As String, ByVal pName As String, ByVal writer As XmlTextWriter)
        writer.WriteStartElement(pID)
        writer.WriteEndElement()

    End Sub

    Private Sub SetNewItems()
        Try
            TOOLS.userField(oCompany, "OINV", "FIRMA ELECTRONICA", 100, "FIRMA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)
            TOOLS.userField(oCompany, "OINV", "CANT. LETRAS", 90, "LETRAS", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)
            TOOLS.userField(oCompany, "OINV", "RESPUESTA", 90, "RESPUESTA", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)
            'TOOLS.userField(oCompany, "OCRD", "NIT", 25, "NIT", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)
            TOOLS.userField(oCompany, "OINV", "ERROR", 100, "ERROR", SAPbobsCOM.BoFieldTypes.db_Alpha, False, SBO_Application)
        Catch ex As Exception
            ex.Message.ToString()
            MessageBox.Show(ex.Message)
        End Try
    End Sub
End Class
