Imports System.Data.OleDb

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        enlace()
        Label1.Text = estado
        conexion.Close()
    End Sub

    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        conexion.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            enlace()
            comando = New OleDb.OleDbCommand("INSERT INTO registros(Nombre, Apellido, Edad)" &
                                 Chr(13) & "VALUES(TextBox1, TextBox2,TextBox3)", conexion)
            comando.Parameters.AddWithValue("@Nombre", TextBox1.Text.ToUpper)
            comando.Parameters.AddWithValue("@Apellido", TextBox2.Text.ToUpper)
            comando.Parameters.AddWithValue("@Edad", TextBox3.Text.ToUpper)
            comando.ExecuteNonQuery()
            MsgBox("Data inserted")
        Catch ex As Exception
            conexion.Close()
            MsgBox("Something went wrong. Data rejected!")
        End Try



    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Read
        Try
            comando = New OleDb.OleDbCommand("SELECT * FROM registros", conexion)
            Dim lector As OleDbDataReader
            Dim tabla As New DataTable()

            conexion.Open()
            lector = comando.ExecuteReader()
            tabla.Load(lector)

            Me.DataGridView1.DataSource = tabla
            conexion.Close()

        Catch ex As Exception
            conexion.Close()
            MsgBox("Something went wrong. Data rejected! <Read>")
        End Try

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'Plot desde el datagridview1

        Chart1.Series.Clear()
        Chart1.Series.Add("Series1")

        With Chart1.Series("Series1")
            .ChartType = DataVisualization.Charting.SeriesChartType.Line
            .MarkerSize = 5
            .Points.Clear()
        End With

        Dim x As String
        Dim y As Integer

        MsgBox(DataGridView1.RowCount - 1)

        For i = 0 To DataGridView1.RowCount - 2
            x = DataGridView1.Rows(i).Cells(1).Value
            y = DataGridView1.Rows(i).Cells(3).Value
            Chart1.Series("Series1").Points.AddXY(x, y)
        Next

    End Sub
End Class
