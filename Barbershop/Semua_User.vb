Imports System.Data.SqlClient
Public Class Semua_User
    Private Sub Semua_User_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Call koneksi()        
        TampilkanData("SELECT * FROM Login", "Login")
        Me.DataGridView1.DataSource = ds.Tables("Login")

        ComboBox1.Items.Add("admin")
        ComboBox1.Items.Add("user")
    End Sub
    Sub TampilkanLagi()
        TampilkanData("SELECT * FROM Login", "Login")
        Me.DataGridView1.DataSource = ds.Tables("Login")
    End Sub
    Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Utama.Show()
        Me.Close()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Transaksi_Terakhir.Show()
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
            TextBox2.Text = .Cells(1).Value
            TextBox3.Text = .Cells(2).Value
            ComboBox1.Text = .Cells(3).Value
        End With
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Call bersih()
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If TextBox1.Text = "" Then
            MsgBox("Silahkan pilih User yang akan dihapus", MsgBoxStyle.Critical, "Kesalahan")
        Else
            SimpanUpdateDelete("DELETE FROM Login WHERE ID_Login = '" & TextBox1.Text & "'")
            Call bersih()
            MsgBox("User berhasil dihapus", MsgBoxStyle.Information, "Informasi")
        End If
        Call TampilkanLagi()
    End Sub
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If TextBox1.Text = "" Then
            MsgBox("Silahkan pilih User yang akan diperbarui", MsgBoxStyle.Critical, "Kesalahan")
        Else
            SimpanUpdateDelete("UPDATE Login SET Username='" & TextBox2.Text & "',Password='" & TextBox3.Text & "',Hak_Akses='" & ComboBox1.Text & "' WHERE ID_Login='" & TextBox1.Text & "'")
            Call bersih()
            MsgBox("User berhasil diperbarui", MsgBoxStyle.Information, "Informasi")
        End If
        Call TampilkanLagi()
    End Sub
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Call ID_Login()
        SimpanUpdateDelete("INSERT INTO Login VALUES ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & ComboBox1.Text & "')")
        Call bersih()
        MsgBox("User baru telah ditambahkan", MsgBoxStyle.Information, "Informasi")
        Call TampilkanLagi()
    End Sub
End Class