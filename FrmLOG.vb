Imports System.Data.SqlClient
Public Class FrmLOG
    Public Fecha As Date
    Private Sub FrmLOG_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim cnnSQL As New SqlConnection()
        Dim cmdSQL_Header As New SqlCommand()

        Dim drsql_Header As SqlDataReader
        Dim StrSQL_Header As String = ""
        Try

            StrSQL_Header = "SELECT" & _
            "RecordKey as RecordKey," & _
            "convert(varchar, Fecha, 103) AS fecha, " & _
            "ErrNum as ErrNum," & _
            "ErrMsj as ErrMsj," & _
            "CASE WHEN (T0.RecordKey<0) THEN (SELECT U_FacCon FROM IMP_ODMK T1 WHERE T1.RecordKey=T0.RecordKey )ELSE" & _
            "(SELECT U_FacOP FROM IMP_ODMK T1 WHERE T1.RecordKey=T0.RecordKey ) END AS Documento" & _
            "FROM LOG T0 WHERE fecha='" + Fecha.ToShortDateString + "' order by RecordKey"

            cnnSQL = New SqlConnection(SQLSERVER_CONNECTIONSTRING)
            cnnSQL.Open()
            cmdSQL_Header = New SqlCommand(StrSQL_Header, cnnSQL)
            drsql_Header = cmdSQL_Header.ExecuteReader()
            Do While drsql_Header.Read()

            Loop

        Catch ex As Exception

        End Try

    End Sub
End Class