Imports System.Data
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.IO
Imports System.Web
Public Class FrmODMK
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ConexionSAP()
        InsertODMK()
    End Sub
    Sub InsertODMK()
        Dim oDocument As SAPbobsCOM.Documents ' Documento de Marketin
        Dim cnnSQL As New SqlConnection()
        Dim cmdSQL_Header As New SqlCommand()
        Dim cmdSQL_Detail As New SqlCommand()
        Dim cmdSQL_LOG As New SqlCommand()
        Dim cmdSQL_SUCESS As New SqlCommand()
        Dim cmdSQL_UPDATE As New SqlCommand()
        Dim cmdSQL_PROCEDURE As New SqlCommand()

        Dim nombre_fecha As Date
        Dim nombre As String        
        nombre_fecha = Today
        nombre = Format(nombre_fecha, "ddMMyyy")
        Dim file_con2 As New FileStream(nombre & ".log", FileMode.Create, FileAccess.Write)
        Dim escritor As New StreamWriter(file_con2)
        Dim TipoDocumento As String = ""

        Dim Fecha_Inicio As String = ""
        Dim Fecha_Fin As String = ""
        Dim sErrMsg As String
        Dim New_sErrMsg As String
        Dim lErrCode As Integer
        Dim ObjType As Integer = 0
        Dim DocNum As String = "IMPORT"


        Try
            Dim drsql_Header As SqlDataReader
            Dim drsql_Detail As SqlDataReader

            Dim StrSQL_Header As String = ""
            Dim StrSQL_Detail As String = ""
            Dim StrSQL_LOG As String = ""
            Dim StrSQL_SUCESS As String = ""
            Dim StrSQL_UPDATE As String = ""

            StrSQL_Header = "SELECT" & _
            " RecordKey" & _
            ",ObjType" & _
            ",CardCode" & _
            ",CardName" & _
            ",Comments" & _
            ",DocCurr" & _
            ",DocDate" & _
            ",U_FacNum" & _
            ",U_FacCon" & _
            ",U_FacOP" & _
            ",U_facturacion" & _
            ",U_FacFecha" & _
            ",U_Sucursal" & _
            ",Series " & _
            ",SlpCode " & _
            ",U_FacReg" & _
            ",U_FacNIT" & _
            ",U_FacGiro" & _
            ",U_FacNom" & _
            ",U_Paquete" & _
            ",U_FacAgencia" & _
            ",U_Comision" & _
            " FROM [Vista_IMP_ODMK] order by [DocDate]"

            cnnSQL = New SqlConnection(SQLSERVER_CONNECTIONSTRING)
            cnnSQL.Open()

            cmdSQL_PROCEDURE.CommandText = "Import_ODMK" ' Stored Procedure to Call
            cmdSQL_PROCEDURE.CommandType = CommandType.StoredProcedure 'Setup Command Type
            cmdSQL_PROCEDURE.Connection = cnnSQL 'Active Connection
            cmdSQL_PROCEDURE.Parameters.AddWithValue("@Refdate", Me.DtpFechaProceso.Value)
            cmdSQL_PROCEDURE.ExecuteNonQuery()

            cmdSQL_Header = New SqlCommand(StrSQL_Header, cnnSQL)
            drsql_Header = cmdSQL_Header.ExecuteReader()
            escritor.WriteLine("RecordKey   Fecha   ErrNum  ErrMsj  TipoDocumento")
            FrmLOG.txtLog.Items.Clear()
            Do While drsql_Header.Read()
                ' ************************************ Aqui se agrega el Encabezado ********************************************
                Select Case drsql_Header.Item("ObjType")
                    Case 13 'facturas
                        oDocument = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                    Case 14 'notas de crédito
                        oDocument = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                    Case 15 'Delivery notes
                        oDocument = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                    Case 16 'Revert delivery notes
                        oDocument = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oReturns)
                    Case 17 'Orders de cliente

                        If drsql_Header.Item("RecordKey") < 0 Then
                            TipoDocumento = drsql_Header.Item("U_FacCon")
                        Else
                            TipoDocumento = drsql_Header.Item("U_FacOP")
                        End If
                        ObjType = 17
                        oDocument = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                        oDocument.CardCode = drsql_Header.Item("CardCode")
                        oDocument.DocDate = CDate(drsql_Header.Item("DocDate")).ToString("yyyy/MM/dd")
                        oDocument.DocDueDate = CDate(drsql_Header.Item("DocDate")).ToString("yyyy/MM/dd")
                        oDocument.Comments = drsql_Header.Item("Comments")
                        oDocument.DocCurrency = drsql_Header.Item("DocCurr")
                        oDocument.Series = drsql_Header.Item("Series").ToString
                        oDocument.SalesPersonCode = drsql_Header.Item("SlpCode").ToString
                        oDocument.UserFields.Fields.Item("U_FacNum").Value = drsql_Header.Item("U_FacNum")
                        oDocument.UserFields.Fields.Item("U_FacCon").Value = drsql_Header.Item("U_FacCon")
                        oDocument.UserFields.Fields.Item("U_FacOP").Value = drsql_Header.Item("U_FacOP")
                        oDocument.UserFields.Fields.Item("U_facturacion").Value = drsql_Header.Item("U_facturacion")
                        oDocument.UserFields.Fields.Item("U_FacFecha").Value = CDate(drsql_Header.Item("DocDate"))
                        oDocument.UserFields.Fields.Item("U_Sucursal").Value = drsql_Header.Item("U_Sucursal")
                        oDocument.UserFields.Fields.Item("U_FacReg").Value = drsql_Header.Item("U_FacReg")
                        oDocument.UserFields.Fields.Item("U_FacNIT").Value = drsql_Header.Item("U_FacNIT")
                        oDocument.UserFields.Fields.Item("U_FacGiro").Value = drsql_Header.Item("U_FacGiro")
                        oDocument.UserFields.Fields.Item("U_FacNom").Value = drsql_Header.Item("U_FacNom")
                        oDocument.UserFields.Fields.Item("U_Paquete").Value = drsql_Header.Item("U_Paquete")
                        oDocument.UserFields.Fields.Item("U_FacAgencia").Value = drsql_Header.Item("U_FacAgencia")
                        oDocument.UserFields.Fields.Item("U_Comision").Value = drsql_Header.Item("U_Comision")
                        oDocument.UserFields.Fields.Item("U_REFACTURACION").Value = drsql_Header.Item("RecordKey").ToString
                    Case 18 'Factura proveedores
                        oDocument = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                    Case 19 'Notas de crédito Proveedores
                        oDocument = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseReturns)
                    Case 20 'Purchase delivery notes
                        oDocument = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
                    Case 22 'Ordenes de Compra
                        oDocument = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                    Case 23 'Cotizaciones
                        oDocument = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQuotations)
                End Select

                StrSQL_Detail = "SELECT " & _
                " RecordKey" & _
                ",LineNum" & _
                ",ItemCode" & _
                ",Dscription" & _
                ",AcctCode" & _
                ",Price" & _
                ",LineTotal" & _
                ",Quantity" & _
                ",UnitPrice" & _
                ",Currency" & _
                ",DiscPrcnt" & _
                ",TaxCode" & _
                ",U_ValorImp" & _
                ",U_contrato" & _
                ",U_fechaIni" & _
                ",U_FechaFin FROM [Vista_IMP_DMK1] WHERE RecordKey=" & drsql_Header.Item(0)

                cmdSQL_Detail = New SqlCommand(StrSQL_Detail, cnnSQL)
                drsql_Detail = cmdSQL_Detail.ExecuteReader()
                Do While drsql_Detail.Read()

                    Select Case drsql_Header.Item("ObjType")
                        Case 17 'Orders de cliente
                            ' ************************************ Aqui se agrega el detalle ********************************************                            
                            oDocument.Lines.ItemCode = drsql_Detail.Item("ItemCode")
                            oDocument.Lines.Price = drsql_Detail.Item("price")
                            oDocument.Lines.LineTotal = drsql_Detail.Item("LineTotal")
                            oDocument.Lines.Quantity = drsql_Detail.Item("Quantity")
                            oDocument.Lines.Currency = drsql_Detail.Item("Currency")
                            oDocument.Lines.DiscountPercent = drsql_Detail.Item("DiscPrcnt")
                            oDocument.Lines.TaxCode = drsql_Detail.Item("TaxCode")
                            oDocument.Lines.UserFields.Fields.Item("U_contrato").Value = drsql_Detail.Item("U_contrato")
                            oDocument.Lines.UserFields.Fields.Item("U_FechIni").Value = drsql_Detail.Item("U_fechaIni")
                            oDocument.Lines.UserFields.Fields.Item("U_FechaFin").Value = drsql_Detail.Item("U_FechaFin")
                            Fecha_Inicio = Format(drsql_Detail.Item("U_fechaIni"), "dd/MM/yyyy")
                            Fecha_Fin = Format(drsql_Detail.Item("U_FechaFin"), "dd/MM/yyyy")
                            ' ************************************ Aqui se agrega el detalle ********************************************
                    End Select
                    oDocument.Lines.Add()
                Loop

                oDocument.SpecialLines.LineType = SAPbobsCOM.BoDocSpecialLineType.dslt_Text
                oDocument.SpecialLines.AfterLineNumber = oDocument.Lines.Count - 2
                If drsql_Header.Item("RecordKey") < 0 Then
                    oDocument.SpecialLines.LineText = "PERIODO DE " & Fecha_Inicio & " AL " & Fecha_Fin
                Else
                    oDocument.SpecialLines.LineText = "IMPRESIÓN DIGITAL"
                End If
                oDocument.SpecialLines.Add()

                oDocument.SpecialLines.LineType = SAPbobsCOM.BoDocSpecialLineType.dslt_Text
                oDocument.SpecialLines.AfterLineNumber = oDocument.Lines.Count - 1
                oDocument.SpecialLines.LineText = "CONTRATO " & drsql_Header.Item("U_FacCon").ToString
                oDocument.SpecialLines.Add()

                drsql_Detail.Close()
                drsql_Detail.Dispose()
                ' ************************************ Aqui se agrega el Encabezado ********************************************
                lErrCode = oDocument.Add()
                sErrMsg = ""
                oCompany.GetLastError(lErrCode, sErrMsg)

                If lErrCode <> 0 Then
                    New_sErrMsg = sErrMsg.Replace("'", "")
                    New_sErrMsg = New_sErrMsg.Replace(",", "")
                    New_sErrMsg = New_sErrMsg.Replace("(", "")
                    New_sErrMsg = New_sErrMsg.Replace(")", "")

                    StrSQL_LOG = "INSERT INTO LOG(ObjType,Fecha,ErrNum,ErrMsj,RecordKey)VALUES(" & ObjType.ToString & ",GETDATE()," & lErrCode.ToString & ",'" & New_sErrMsg & "'," & drsql_Header.Item("RecordKey").ToString & ")"
                    cmdSQL_LOG = New SqlCommand(StrSQL_LOG, cnnSQL)
                    cmdSQL_LOG.ExecuteNonQuery()
                    escritor.WriteLine(drsql_Header.Item("RecordKey").ToString & "  " & nombre & "  " & lErrCode.ToString & "   " & New_sErrMsg & "   " + TipoDocumento)
                    FrmLOG.txtLog.Items.Add(drsql_Header.Item("RecordKey").ToString & "  " & nombre & "  " & lErrCode.ToString & "   " & New_sErrMsg + "  " + TipoDocumento)
                Else
                    StrSQL_SUCESS = "INSERT INTO IMP_ODMK_IMPORTED(RecordKey,DocNum,ObjType,ProcessDate)VALUES(" & drsql_Header.Item("RecordKey").ToString & ",'" & DocNum & "'," & ObjType.ToString & ", GETDATE())"
                    cmdSQL_SUCESS = New SqlCommand(StrSQL_SUCESS, cnnSQL)
                    cmdSQL_SUCESS.ExecuteNonQuery()
                    escritor.WriteLine(drsql_Header.Item("RecordKey").ToString & "  " & nombre & "  " & "0" & "   " & "Sin Error" & "   " & TipoDocumento)
                End If

            Loop
            drsql_Header.Close()
            cnnSQL.Close()
            drsql_Header.Dispose()
            cnnSQL.Dispose()
        Catch err As Exception
            MsgBox(err.Message)
        End Try        
        escritor.Close()
        MsgBox("Importación Finalizada")
        DesConexionSAP()
        FrmLOG.ShowDialog()
    End Sub

End Class