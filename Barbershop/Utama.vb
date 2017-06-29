Imports System.Data.SqlClient
Public Class Utama
    Private Sub Utama_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call koneksi()
        Call ID_Transaksi()

        'Tampilkan Layanan ke ComboBox1
        DataTable("SELECT ID_Layanan,Nama_Layanan FROM Layanan")
        Me.ComboBox1.DataSource = dt
        Me.ComboBox1.ValueMember = "ID_Layanan"
        Me.ComboBox1.DisplayMember = "Nama_Layanan"

        'Tampilan Nama Pelanggan Secara Otomatis ke TextBox2
        DataTable("SELECT Nama_Pelanggan FROM Pelanggan")
        col = New AutoCompleteStringCollection
        For i As Integer = 0 To dt.Rows.Count - 1
            col.Add(dt.Rows(i)("Nama_Pelanggan"))
        Next
        TextBox2.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox2.AutoCompleteCustomSource = col
        TextBox2.AutoCompleteMode = AutoCompleteMode.Suggest

        Debug.Print(con.State)
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Transaksi_Terakhir.Show()
        'Me.Close()
        Me.Hide()
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Semua_User.Show()
        Me.Hide()
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Layanan_Pelanggan.Show()
        Me.Hide()
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Data_Pelanggan.Show()
        Me.Hide()
    End Sub
    Public Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Call ID_Transaksi()

        Dim Total As Integer
        Dim Baris As Integer

        DataGridView1.Rows.Add(1)
        DataGridView1.Rows(DataGridView1.Rows.Count - 2).Cells(0).Value = CStr(DataGridView1.RowCount - 1)
        DataGridView1.Rows(DataGridView1.Rows.Count - 2).Cells(1).Value = DateTimePicker1.Text
        DataGridView1.Rows(DataGridView1.Rows.Count - 2).Cells(2).Value = TextBox2.Text
        DataGridView1.Rows(DataGridView1.Rows.Count - 2).Cells(3).Value = ComboBox1.Text
        DataGridView1.Rows(DataGridView1.Rows.Count - 2).Cells(4).Value = TextBox3.Text
        DataGridView1.Update()

        For Baris = 0 To DataGridView1.RowCount - 1
            Total = Total + DataGridView1.Rows(Baris).Cells(4).Value
        Next
        TextBox4.Text = Total
        SimpanUpdateDelete("INSERT INTO Transaksi VALUES ('" & TextBox1.Text & "','" & DateTimePicker1.Text & "','" & TextBox2.Text & "','" & ComboBox1.Text & "','" & TextBox4.Text & "')")
        Call ID_Transaksi()
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim Kembalian As String
        Kembalian = Val(TextBox5.Text) - Val(TextBox4.Text)
        TextBox6.Text = Kembalian
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        cmd = New SqlCommand("SELECT Harga_Layanan FROM Layanan WHERE Nama_Layanan='" & ComboBox1.Text & "'", con)
        rd = cmd.ExecuteReader()
        rd.Read()
        If rd.HasRows Then
            TextBox3.Text = rd.Item("Harga_Layanan")
        End If
        rd.Close()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Me.Close()
    End Sub
End Class