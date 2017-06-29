Imports System.Data.SqlClient
Public Class Data_Pelanggan
    Private Sub Pelanggan_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TampilkanData("SELECT * FROM Pelanggan", "Pelanggan")
        Me.DataGridView1.DataSource = ds.Tables("Pelanggan")
    End Sub
    Sub TampilkanLagi()
        TampilkanData("SELECT * FROM Pelanggan", "Pelanggan")
        Me.DataGridView1.DataSource = ds.Tables("Pelanggan")
    End Sub
    Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        RichTextBox1.Text = ""
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Utama.Show()
        Me.Close()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Transaksi_Terakhir.Show()
        Me.Close()
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Semua_User.Show()
        Me.Close()
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Layanan_Pelanggan.Show()
        Me.Close()
    End Sub
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim i As Integer
        i = Me.DataGridView1.CurrentRow.Index
        With DataGridView1.Rows.Item(i)
            TextBox1.Text = .Cells(0).Value
            TextBox2.Text = .Cells(1).Value
            TextBox3.Text = .Cells(2).Value
            RichTextBox1.Text = .Cells(3).Value
        End With
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Call bersih()
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If TextBox1.Text = "" Then
            MsgBox("Silahkan pilih Pelanggan yang akan dihapus", MsgBoxStyle.Critical, "Kesalahan")
        Else
            SimpanUpdateDelete("DELETE FROM Pelanggan WHERE ID_Pelanggan = '" & TextBox1.Text & "'")
            Call bersih()
            MsgBox("Pelanggan berhasil dihapus", MsgBoxStyle.Information, "Informasi")
        End If
        Call TampilkanLagi()
    End Sub
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If TextBox1.Text = "" Then
            MsgBox("Silahkan pilih Pelanggan yang akan diperbarui", MsgBoxStyle.Critical, "Kesalahan")
        Else
            SimpanUpdateDelete("UPDATE Pelanggan SET Nama_Pelanggan='" & TextBox2.Text & "', Kontak_Pelanggan='" & TextBox3.Text & "', Alamat_Pelanggan='" & RichTextBox1.Text & "' WHERE ID_Pelanggan='" & TextBox1.Text & "'")
            Call bersih()
            MsgBox("Data Pelanggan berhasil diperbarui", MsgBoxStyle.Information, "Informasi")
        End If
        Call TampilkanLagi()
    End Sub
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Call ID_Pelanggan()
        SimpanUpdateDelete("INSERT INTO Pelanggan VALUES ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & RichTextBox1.Text & "')")
        Call bersih()
        MsgBox("Pelanggan baru telah ditambahkan", MsgBoxStyle.Information, "Informasi")
        Call TampilkanLagi()
    End Sub
End Class