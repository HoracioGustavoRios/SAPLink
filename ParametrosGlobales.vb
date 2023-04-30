Imports System.Data
Imports System.Data.SqlClient
Module ParametrosGlobales
    Public ruta_INI As String
    Public HASHKEY As String = "key"
    Public SQLSERVER_CONNECTIONSTRING As String = "data source={ServerName};initial catalog={DataBase};password={Password};persist security info=True;user id=sa;workstation id=horacio;MultipleActiveResultSets=True"
    Public oCompany As SAPbobsCOM.Company
    Public conectado_a_sap As Boolean

    Sub SysRutaINI()
        ruta_INI = Application.StartupPath & "\CONFIG.INI"
    End Sub
    Function SysProbar_conexion(ByVal conexion As String) As Boolean
        Dim IsConnecting As Boolean = True
        Dim cnnSQL As SqlConnection
        Dim cmdSQL As SqlCommand
        Dim strSQL As String
        While IsConnecting
            Try
                Dim drsql As SqlDataReader
                strSQL = "select min(name) from sysobjects"
                cnnSQL = New SqlConnection(conexion)
                cnnSQL.Open()
                cmdSQL = New SqlCommand(strSQL, cnnSQL)
                drsql = cmdSQL.ExecuteReader()
                IsConnecting = False
                SysProbar_conexion = True
                drsql.Close()
                cmdSQL.Dispose()
                cnnSQL.Close()
                cnnSQL.Dispose()
            Catch
                SysProbar_conexion = False
                IsConnecting = False
            End Try
        End While
    End Function

    Sub ConexionSAP()
        Dim myArchivoINI As New cIniArray
        Dim lRetCode, lErrCode As Long
        Dim sErrMsg As String = ""
        SysRutaINI()
        oCompany = New SAPbobsCOM.Company
        oCompany.UseTrusted = False
        oCompany.Server = myArchivoINI.IniGet(ParametrosGlobales.ruta_INI, "SAP", "SERVER_NAME", "")
        oCompany.CompanyDB = myArchivoINI.IniGet(ParametrosGlobales.ruta_INI, "LINE_COMMAND_PARAMETERS", "", "")
        oCompany.DbUserName = myArchivoINI.IniGet(ParametrosGlobales.ruta_INI, "LINE_COMMAND_PARAMETERS", "SQL_USER", "")
        oCompany.DbPassword = PrivateKey.Decrypt(myArchivoINI.IniGet(ParametrosGlobales.ruta_INI, "LINE_COMMAND_PARAMETERS", "SQL_PASS", ""), ParametrosGlobales.HASHKEY)
        oCompany.UserName = myArchivoINI.IniGet(ParametrosGlobales.ruta_INI, "LINE_COMMAND_PARAMETERS", "SAP_USER", "")
        oCompany.Password = PrivateKey.Decrypt(myArchivoINI.IniGet(ParametrosGlobales.ruta_INI, "LINE_COMMAND_PARAMETERS", "SAP_PASS", ""), ParametrosGlobales.HASHKEY)
        oCompany.language = SAPbobsCOM.BoSuppLangs.ln_Spanish_La
        oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
        oCompany.LicenseServer = myArchivoINI.IniGet(ParametrosGlobales.ruta_INI, "SAP", "LICENSE_SERVER", "") & ":30000"

        lRetCode = oCompany.Connect
        If oCompany.Connected = True Then
            conectado_a_sap = True
            Exit Sub
        End If
        If lRetCode <> 0 Then
            oCompany.GetLastError(lErrCode, sErrMsg)
            MsgBox(sErrMsg)
            conectado_a_sap = False
        Else
            conectado_a_sap = True
        End If
    End Sub
    Sub DesConexionSAP()
        oCompany.Disconnect()
    End Sub
End Module
