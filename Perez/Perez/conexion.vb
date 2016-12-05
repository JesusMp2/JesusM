Imports System.Data.Sql
Imports System.Data.SqlClient


Module Conexion
    Public conexiones As SqlConnection
    Public enunciado As SqlCommand
    Public respuesta As SqlDataReader
    Public adaptador As SqlDataAdapter

    'FUNCION DE CONEXION A LA BASE DE DATOS' 

    '***********************************­****************************************­****************************************­****************************' 

    Sub abrir()
        Try
            conexiones = New SqlConnection("Data Source=JESUS\JESUSM;Initial Catalog=Tutorial;Integrated Security=True")
            conexiones.Open()
            MsgBox("Conexion exitosa", MsgBoxStyle.Information, "Se ha conectado correctamente") '
        Catch ex As Exception
            MsgBox("Error al realizar la conexion" & ex.Message, MsgBoxStyle.Critical, "Error de conexion")
            conexiones.Close() 'Cierra la conexion' 
        End Try
    End Sub
    '***********************************­****************************************­****************************************­****************************' 
    'FUNCION DE INSERCION ' 
    '***********************************­****************************************­****************************************­********************************' 
    Function insercionPersona(ByVal identificacion As String, ByVal nombre As String, ByVal fecha As Date, ByVal edad As Integer) As String

        Dim resultado As String = ""
        Try
            enunciado = New SqlCommand("insert into Persona(identificacion,nombre,fecha,edad) values('" & identificacion & "','" & nombre & "','" & fecha & "'," & edad & ")", conexiones)
            enunciado.ExecuteNonQuery()
            resultado = "Se inserto correctamente"
            conexiones.Close()

        Catch ex As Exception
            resultado = "No se inserto" + ex.ToString
            conexiones.Close()

        End Try
        Return resultado
    End Function
    '***********************************­****************************************­****************************************­***********************************'
    Function actualizarNombre(ByVal identificacion As String, ByVal nuevo As String) As String
        Dim resultado As String = ""
        Try
            enunciado = New SqlCommand("Update Persona set Nombre= '" & nuevo & "' where Identificacion='" & identificacion & "'", conexiones)
            enunciado.ExecuteNonQuery()
            resultado = "Se Actualizo Correctamente los datos"
        Catch ex As Exception
            resultado = "No se actualizo los datos correctamente" + ex.ToString
        End Try
        Return resultado
    End Function
End Module