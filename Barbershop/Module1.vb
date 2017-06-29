Imports System.Data.SqlClient
Module Module1
    Public con As SqlConnection
    Public col As AutoCompleteStringCollection
    Public cmd As SqlCommand
    Public ds As DataSet
    Public dr As DataRow
    Public dt As DataTable
    Public da As SqlDataAdapter
    Public rd As SqlDataReader
    Public db As String
    Public Sub koneksi()
        'db = "Integrated Security=SSPI; Persist Security Info=True; Initial Catalog=ardi; Data Source=(local)"
        db = "Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=Barbershop;Integrated Security=True;Persist Security Info=True"
        con = New SqlConnection(db)
        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
                'MsgBox("Koneksi Sukses")
            End If
        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Kesalahan")
        End Try
    End Sub
    Public Function SQLTable(ByVal Source As String) As DataTable
        Try
            da = New SqlDataAdapter(Source, con)
            dt = New DataTable
            da.Fill(dt)
            SQLTable = dt
        Catch ex As SqlClient.SqlException
            MsgBox(ex.Message)
            SQLTable = Nothing
        End Try
    End Function
    Public Sub ID_Pelanggan()
        Dim s As String
        dr = SQLTable("SELECT MAX(RIGHT(ID_Pelanggan,1)) AS Nomor FROM Pelanggan").Rows(0)
        If dr.IsNull("Nomor") Then
            s = "MBR-1"
        Else
            s = "MBR-" & Format(dr("Nomor") + 1, "0")
        End If
        Data_Pelanggan.TextBox1.Text = s
    End Sub
    Public Sub ID_Layanan()
        Dim s As String
        dr = SQLTable("SELECT MAX(RIGHT(ID_Layanan,1)) AS Nomor FROM Layanan").Rows(0)
        If dr.IsNull("Nomor") Then
            s = "SVC-1"
        Else
            s = "SVC-" & Format(dr("Nomor") + 1, "0")
        End If
        Layanan_Pelanggan.TextBox1.Text = s
    End Sub
    Public Sub ID_Transaksi()
        Dim s As String
        dr = SQLTable("SELECT MAX(RIGHT(ID_Transaksi,1)) AS Nomor FROM Transaksi").Rows(0)
        If dr.IsNull("Nomor") Then
            s = "TX-1"
        Else
            s = "TX-" & Format(dr("Nomor") + 1, "0")
        End If
        Utama.TextBox1.Text = s
        Utama.TextBox1.Enabled = False
    End Sub
    Public Sub ID_Login()
        Dim s As String
        dr = SQLTable("SELECT MAX(RIGHT(ID_Login,1)) AS Nomor FROM Login").Rows(0)
        If dr.IsNull("Nomor") Then
            s = "USR-1"
        Else
            s = "USR-" & Format(dr("Nomor") + 1, "0")
        End If
        Semua_User.TextBox1.Text = s
    End Sub
    Public Function Masuk(ByVal query As String)
        Try
            cmd = New SqlCommand(query, con)
            rd = cmd.ExecuteReader()
            rd.Read()
            If rd.HasRows Then
                Login.Hide()
                If rd("Hak_Akses") = "admin" Then
                    Utama.Show()
                Else
                    Utama.Show()
                    Utama.Button2.Enabled = False
                End If
            Else
                MsgBox("Username dan Password tidak ditemukan", MsgBoxStyle.Critical, "Kesalahan")
            End If
            rd.Close()
        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Kesalahan")
        End Try
    End Function
    Public Function DataTable(ByVal query As String)
        Try
            da = New SqlDataAdapter(query, con)
            dt = New DataTable
            da.Fill(dt)
        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Kesalahan")
        End Try
    End Function
    Public Function SimpanUpdateDelete(ByVal query As String)
        Try
            cmd = New SqlCommand(query, con)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Kesalahan")
        End Try
    End Function
    Public Function TampilkanData(ByVal query As String, ByVal nama_table As String)
        Try
            da = New SqlDataAdapter(query, con)
            ds = New DataSet
            ds.Clear()
            da.Fill(ds, nama_table)
        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Kesalahan")
        End Try
    End Function
End Module