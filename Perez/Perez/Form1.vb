Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MsgBox(insercionPersona(TextBox1.Text, TextBox2.Text, CDate(Time1.Text), CInt(TextBox3.Text)))
        limpiar()
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        MsgBox(actualizarNombre(TextBox4.Text, TextBox5.Text))
        limpiar()
    End Sub

    Sub limpiar()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()

    End Sub

End Class
