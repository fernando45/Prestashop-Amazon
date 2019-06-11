Imports System.IO
Imports System.Net
Imports System.Xml
Public Class Form1
    Const Clave As String = "C23D8TF675W562UCJDL27J3UU4AMR4V8"
    Const DirApi As String = "https://www.topkit.es/api/products"
    Private Productos As New ClsProducts
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Productos.Construye()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Close()
    End Sub
End Class

