Imports System.Data.SqlClient
Public Class Layanan_Pelanggan
    Private Sub Layanan_Pelanggan_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TampilkanData("SELECT * FROM Layanan", "Layanan")
        Me.DataGridView1.DataSource = ds.Tables("Layanan")
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
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Data_Pelanggan.Show()
        Me.Close()
    End Sub
    Sub TampilanLagi()
        TampilkanData("SELECT * FROM Layanan", "Layanan")
        Me.DataGridView1.DataSource = ds.Tables("Layanan")
    End Sub
    Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If TextBox1.Text = "" Then
            MsgBox("Silahkan pilih Layanan yang akan dihapus", MsgBoxStyle.Critical, "Kesalahan")
        Else
            SimpanUpdateDelete("DELETE FROM Layanan WHERE ID_Layanan = '" & TextBox1.Text & "'")
            Call bersih()
            MsgBox("Layanan berhasil dihapus", MsgBoxStyle.Information, "Informasi")
        End If
        Call TampilanLagi()
    End Sub
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If TextBox1.Text = "" Then
            MsgBox("Silahkan pilih Layanan yang akan diperbarui", MsgBoxStyle.Critical, "Kesalahan")
        Else
            SimpanUpdateDelete("UPDATE Layanan SET Nama_Layanan = '" & TextBox2.Text & "', Harga_Layanan = '" & TextBox3.Text & "' WHERE ID_Layanan = '" & TextBox1.Text & "'")
            Call bersih()
            MsgBox("Data Layanan berhasil diperbarui", MsgBoxStyle.Information, "Informasi")
        End If
        Call TampilanLagi()
    End Sub
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Call ID_Layanan()
        SimpanUpdateDelete("INSERT INTO Layanan VALUES ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "')")
        Call bersih()
        MsgBox("Layanan baru telah ditambahkan", MsgBoxStyle.Information, "Informasi")
        Call TampilanLagi()
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Call bersih()
    End Sub
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim i As Integer
        i = Me.DataGridView1.CurrentRow.Index
        With DataGridView1.Rows.Item(i)
            TextBox1.Text = .Cells(0).Value
            TextBox2.Text = .Cells(1).Value
            TextBox3.Text = .Cells(2).Value
        End With
    End Sub
End Class