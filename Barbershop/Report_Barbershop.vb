Imports CrystalDecisions.CrystalReports.Engine
Public Class Report_Barbershop
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim cryrpt As New ReportDocument
        cryrpt.Load("Transaksi_Barbershop.rpt")
        Me.CrystalReportViewer1.ReportSource = cryrpt
        Me.CrystalReportViewer1.Refresh()
    End Sub
End Class