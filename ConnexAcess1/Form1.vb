﻿Imports System.ComponentModel
Imports System.Data.OleDb

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        enlace()
        Label1.Text = estado
        TextBox4.Text = Environment.UserName
        TextBox6.Text = Today()

        cargar_lista_lanes()
        cargar_lista_motores()

        conexion.Close()

        ComboBox1.SelectedIndex = ComboBox1.FindStringExact("Pack Lane 1") 'MTR 61302


        read_Readings_table()
        plot_readings()
    End Sub

    Sub cargar_lista_lanes()
        With ComboBox1.Items
            For i = 1 To 11
                .Add("Pack Lane " & i)
            Next
            For i = 1 To 4
                .Add("MOD Lane " & i)
            Next
            .Add("CP31")
            .Add("CP32")
            .Add("LOOP 1 of 2")
            .Add("LOOP 2 of 2")
            .Add("Lanes 1/2")
            .Add("Lanes 2/2")
            .Add("Tranship")
        End With
    End Sub

    Sub cargar_lista_motores()

    End Sub

    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        conexion.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        plot_readings()
    End Sub

    Sub plot_readings()
        'Plot desde el datagridview2

        Chart1.Series.Clear()
        Chart1.Series.Add("NDE")
        Chart1.Series.Add("DE")
        Chart1.Series.Add("GOB")

        'Chart1.Legends.Clear()
        Chart1.ChartAreas(0).AxisX.Interval = 1
        Chart1.ChartAreas(0).AxisY.MajorGrid.LineColor = Color.FromArgb(&H50, &H9C, &H9A, &H95)
        'Chart1.ChartAreas(0).AxisX.MajorGrid.LineColor = Color.FromArgb(&H50, &H9C, &H9A, &H95)
        Chart1.ChartAreas(0).AxisX.MajorGrid.Enabled = False
        Chart1.ChartAreas(0).AxisX.LabelStyle.Enabled = False
        Chart1.ChartAreas(0).AxisX.MajorTickMark.Enabled = False

        With Chart1.Series("NDE")
            .ChartType = DataVisualization.Charting.SeriesChartType.Spline
            '.MarkerStyle = DataVisualization.Charting.MarkerStyle.Circle
            .MarkerSize = 5
            .Points.Clear()
        End With

        With Chart1.Series("DE")
            .ChartType = DataVisualization.Charting.SeriesChartType.Spline
            '.MarkerStyle = DataVisualization.Charting.MarkerStyle.Circle
            .MarkerSize = 5
            .Points.Clear()
        End With

        With Chart1.Series("GOB")
            .ChartType = DataVisualization.Charting.SeriesChartType.Spline
            '.MarkerStyle = DataVisualization.Charting.MarkerStyle.Circle
            .MarkerSize = 5
            .Points.Clear()
        End With

        Dim x As String
        Dim y As Integer

        For i = 0 To DataGridView2.RowCount - 2
            x = i 'DataGridView2.Rows(i).Cells(0).Value
            y = DataGridView2.Rows(i).Cells(6).Value
            If y <> 0 Then
                Chart1.Series("NDE").Points.AddXY(x, y)
            End If
        Next

        For i = 0 To DataGridView2.RowCount - 2
            x = i 'DataGridView2.Rows(i).Cells(0).Value
            y = DataGridView2.Rows(i).Cells(7).Value
            If y <> 0 Then
                Chart1.Series("DE").Points.AddXY(x, y)
            End If
        Next

        For i = 0 To DataGridView2.RowCount - 2
            x = i 'DataGridView2.Rows(i).Cells(0).Value
            y = DataGridView2.Rows(i).Cells(9).Value
            If y <> 0 Then
                Chart1.Series("GOB").Points.AddXY(x, y)
            End If
        Next


    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'OK: Ingresar los registros a la tabla Motores
        Try

            enlace()

            comando = New OleDb.OleDbCommand("INSERT INTO Readings(MTR, Tech, DTime, WO, Lane, DE, NDE, GIB, GOB)" & Chr(13) _
                                             & "VALUES(TextBox9, TextBox4, TextBox6, TextBox5, TextBox7, TextBox10, TextBox8, TextBox11, TextBox12)", conexion)

            comando.Parameters.AddWithValue("@MTR", TextBox9.Text)
            comando.Parameters.AddWithValue("@Tech", TextBox4.Text)
            comando.Parameters.AddWithValue("@DTime", TextBox6.Text)
            comando.Parameters.AddWithValue("@WO", TextBox5.Text)
            comando.Parameters.AddWithValue("@Lane", TextBox7.Text)
            comando.Parameters.AddWithValue("@DE", Val(TextBox10.Text))
            comando.Parameters.AddWithValue("@NDE", Val(TextBox8.Text))
            comando.Parameters.AddWithValue("@GIB", Val(TextBox11.Text))
            comando.Parameters.AddWithValue("@GOB", Val(TextBox12.Text))

            comando.ExecuteNonQuery()
            MsgBox("Data inserted by " & Environment.UserName)
            conexion.Close()

            read_Readings_table()

        Catch ex As Exception
            conexion.Close()
            MsgBox("Something went wrong. Data rejected!")
        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        read_Readings_table()
        plot_readings()
    End Sub

    Sub read_Readings_table()
        'Read from Readings
        Try
            Dim MTR As String
            MTR = TextBox9.Text

            'comando = New OleDb.OleDbCommand("SELECT * FROM Readings", conexion)

            comando = New OleDb.OleDbCommand("SELECT * FROM Readings WHERE MTR='" & MTR & "'", conexion)

            Dim lector As OleDbDataReader
            Dim tabla As New DataTable()

            conexion.Open()
            lector = comando.ExecuteReader()
            tabla.Load(lector)

            Me.DataGridView2.DataSource = tabla
            DataGridView2.Sort(DataGridView2.Columns(4), ListSortDirection.Descending) 'Organizar por WO descendiente

            conexion.Close()

        Catch ex As Exception
            conexion.Close()
            MsgBox("Something went wrong. Data rejected! <Read>")
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        'Read MTR from Motors
        Try

            Dim Lane As String
            Lane = ComboBox1.Text
            TextBox7.Text = Lane

            comando = New OleDb.OleDbCommand("SELECT MTR FROM Motors_List WHERE Lane='" & Lane & "'", conexion)

            Dim lector As OleDbDataReader
            Dim tabla As New DataTable()

            conexion.Open()
            lector = comando.ExecuteReader()
            tabla.Load(lector)

            Me.DataGridView3.DataSource = tabla
            DataGridView3.ReadOnly = True
            DataGridView3.ColumnHeadersVisible = False
            DataGridView3.RowHeadersVisible = False
            DataGridView3.AllowUserToAddRows = False
            DataGridView3.AllowUserToDeleteRows = False
            DataGridView3.AllowUserToResizeRows = False
            DataGridView3.AllowUserToResizeColumns = False
            DataGridView3.AllowUserToOrderColumns = False
            DataGridView3.Columns(0).Width = DataGridView3.Width

            conexion.Close()

        Catch ex As Exception
            conexion.Close()
            MsgBox("Something went wrong! <Read Motors_List>")
        End Try
    End Sub

    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick
        TextBox9.Text = DataGridView3.CurrentCell.Value
        read_Readings_table()
        plot_readings()

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Try
            Dim t As Integer
            t = Val(TextBox13.Text)
            read_Readings_table()
            truncar_lista(t)
            plot_readings()
        Catch ex As Exception
            MsgBox("This operation produced an error. Is this error persists contact administrator.")
        End Try
    End Sub

    Sub truncar_lista(r As Integer)
        Dim ufil As Integer
        ufil = DataGridView2.Rows.Count - 2
        For i = ufil To r Step -1
            DataGridView2.Rows.RemoveAt(i)
        Next
    End Sub

End Class
