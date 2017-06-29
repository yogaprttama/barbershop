'Dibuat oleh Yoga Pratama
'Pada 28/Mei/2017
'Diperbarui pada 2/Agustus/2017
'Sumber :
'

Imports System.Data.SqlClient
Public Class Login
    Private Sub Login_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call koneksi()
    End Sub
    'Label X
    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        Me.Close()
    End Sub
    'Tombol Login
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Masuk("SELECT ID_Login,Hak_akses FROM Login WHERE Username='" & TextBox1.Text & "' AND Password='" & TextBox2.Text & "'")
    End Sub
    'Tampilkan Password
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            TextBox2.UseSystemPasswordChar = False
        Else
            TextBox2.UseSystemPasswordChar = True
        End If
    End Sub
End Class