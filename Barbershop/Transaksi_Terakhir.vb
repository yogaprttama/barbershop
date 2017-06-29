Public Class Transaksi_Terakhir
    Sub Tampilkan()
        'Tampilkan Data Transaksi        
        TampilkanData("SELECT * FROM Transaksi", "Transaksi")
        Me.DataGridView1.DataSource = ds.Tables("Transaksi")
        Me.DataGridView1.Refresh()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Utama.Show()
        'Me.Close()
        Me.Hide()
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Semua_User.Show()
        Me.Close()
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Layanan_Pelanggan.Show()
        Me.Close()
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Data_Pelanggan.Show()
        Me.Close()
    End Sub
    Private Sub DataGridView1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        Dim i As Integer
        i = Me.DataGridView1.CurrentRow.Index
        With DataGridView1.Rows.Item(i)
            TextBox1.Text = .Cells(0).Value
            TextBox2.Text = .Cells(2).Value
        End With
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If TextBox1.Text = "" Then
            MsgBox("Silahkan pilih Transaksi yang akan dihapus", MsgBoxStyle.Critical, "Kesalahan")
        Else
            'Hapus Transaksi
            SimpanUpdateDelete("DELETE FROM Transaksi WHERE ID_Transaksi = '" & TextBox1.Text & "'")
            MsgBox("Transaksi berhasil dihapus", MsgBoxStyle.Information, "Informasi")
        End If

        'Bersihkan TextBox1 dan TextBox2
        TextBox1.Text = ""
        TextBox2.Text = ""

        'Tampilkan Data Kembali
        Call Tampilkan()
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Report_Barbershop.Show()
    End Sub
    Private Sub Transaksi_Terakhir_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call Tampilkan()
    End Sub
End Class