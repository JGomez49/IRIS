
Imports System.Data
Imports System.Data.OleDb

Module Module1
    Public conexion As New OleDbConnection
    Public estado As String
    Public comando As New OleDbCommand

    Sub enlace()
        Try
            'conexion.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\ajhgonz\Documents\registro.accdb"
            '\\ant\dept-na\YYC1\Support\Facilities\DBIR\registro.accdb
            conexion.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\ant\dept-na\YYC1\Support\Facilities\DBIR\registro.accdb"
            conexion.Open()
            estado = "Status: Connected"

        Catch ex As Exception
            estado = "Status: Not connected"

        End Try
    End Sub
End Module
