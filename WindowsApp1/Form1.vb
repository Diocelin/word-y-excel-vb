Public Class Form1
    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Cerrar()
    End Sub

    Private Sub BCrear_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Crear()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles BCrear.Click
        Guardarcomo()
    End Sub
End Class
