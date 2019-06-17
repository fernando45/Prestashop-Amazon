Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Text.Encoding
Imports System.Threading
Public Class Form1

    Private Productos As New ClsProducts

    Public cts As CancellationTokenSource
    Public WithEvents work As New System.ComponentModel.BackgroundWorker

    Const destino As String = "ftp://ftp.cluster011.hosting.ovh.net/amazon"
    Const user As String = "docshareuz-topkit"
    Const pass As String = "i5OEhgoe7gpVgKQ"
    Private fichero As String = Application.StartupPath & "\productsE.xml"

    ' Public token As CancellationToken = cts.Token


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.work.WorkerSupportsCancellation = True
    End Sub

    Private Sub Work_DoWork(sender As System.Object, e As System.ComponentModel.DoWorkEventArgs) Handles work.DoWork

        Productos.Construye()

    End Sub
    Private Sub BackgroundWorker_ProgressChanged(sender As System.Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles work.ProgressChanged

        '   Me.ProgressBar1.Value += e.ProgressPercentage

    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As System.Object, e As System.ComponentModel.DoWorkEventArgs) Handles work.DoWork

        Do
            If work.CancellationPending Then
                e.Cancel = True
                Exit Do
            Else

                Exit Do
            End If
        Loop

    End Sub
    Private Sub BackgroundWorker_RunWorkerCompleted(sender As System.Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles work.RunWorkerCompleted

        If e.Cancelled Then
            MessageBox.Show("Laación ha sido cancelada.")
            Me.Panel1.Visible = False
            Me.Button1.Enabled = True

        ElseIf e.Error IsNot Nothing Then
            ' MessageBox.Show("Seroducido un error durante la ejecución: " & e.Error.Message)
        Else
            Dim uri As String
            uri = destino & "/products.xml"
            My.Computer.Network.UploadFile(fichero, uri, user, pass)
            Me.Panel1.Visible = False
            Me.Button1.Enabled = True

            MessageBox.Show("La operacion ha finalizado con éxito.")
        End If



    End Sub
    Public Function subirFichero() As String

        Try
            ' My.Computer.Network.UploadFile(fichero, "Servidor FTP y archivos", "Nombre de usuario", "contraseña")

            Dim infoFichero As New FileInfo(fichero)

            Dim uri As String
            uri = destino & "/products.xml"
            ' My.Computer.Network.UploadFile(fichero, uri, user, pass)

            '' Si no existe el directorio, lo creamos
            'If Not existeObjeto(dir, user2, pass2) Then
            '    creaDirectorio(dir)
            'End If

            Dim peticionFTP As FtpWebRequest

            ' Creamos una peticion FTP con la dirección del fichero que vamos a subir
            peticionFTP = CType(FtpWebRequest.Create(New Uri(destino)), FtpWebRequest)

            ' Fijamos el usuario y la contraseña de la petición
            peticionFTP.Credentials = New NetworkCredential(user, pass)

            peticionFTP.KeepAlive = False
            peticionFTP.UsePassive = False

            ' Seleccionamos el comando que vamos a utilizar: Subir un fichero
            peticionFTP.Method = WebRequestMethods.Ftp.UploadFile

            ' Especificamos el tipo de transferencia de datos
            peticionFTP.UseBinary = True

            ' Informamos al servidor sobre el tamaño del fichero que vamos a subir
            peticionFTP.ContentLength = infoFichero.Length

            ' Fijamos un buffer de 2KB
            Dim longitudBuffer As Integer
            longitudBuffer = 2048
            Dim lector As Byte() = New Byte(2048) {}

            Dim num As Integer

            ' Abrimos el fichero para subirlo
            Dim fs As FileStream
            fs = infoFichero.OpenRead()

            Dim escritor As Stream
            escritor = peticionFTP.GetRequestStream()

            ' Leemos 2 KB del fichero en cada iteración
            num = fs.Read(lector, 0, longitudBuffer)

            While (num <> 0)
                ' Escribimos el contenido del flujo de lectura en el 
                ' flujo de escritura del comando FTP
                escritor.Write(lector, 0, num)
                num = fs.Read(lector, 0, longitudBuffer)
            End While

            escritor.Close()
            fs.Close()
            peticionFTP = Nothing
            ' Si todo ha ido bien, se devolverá String.Empty
            Return String.Empty

        Catch ex1 As WebException
            Dim ole As String = ex1.Message
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            Return ex.Message
        End Try
        Return String.Empty
    End Function
    Public Sub DescargaStock()
        Dim Ruta As String = Application.StartupPath
        Ruta = Ruta & "\"
        Dim ficFTP As String = "ftp://ftp.cluster011.hosting.ovh.net/DGLexistencias.csv"
        Try

            'Eliminar_Archivo(Ruta & "DGLexistencias.csv")

            Dim dirFtp As FtpWebRequest = CType(FtpWebRequest.Create(ficFTP), FtpWebRequest)

            ' Los datos del usuario (credenciales)
            Dim cr As New NetworkCredential("docshareuz-topkit", "i5OEhgoe7gpVgKQ")
            dirFtp.Credentials = cr

            ' El comando a ejecutar usando la enumeración de WebRequestMethods.Ftp
            dirFtp.Method = WebRequestMethods.Ftp.DownloadFile

            ' Obtener el resultado del comando
            Dim reader As New StreamReader(dirFtp.GetResponse().GetResponseStream())

            ' Leer el stream (el contenido del archivo)
            Dim res As String = reader.ReadToEnd()

            ' Mostrarlo.
            'Console.WriteLine(res)

            ' Guardarlo localmente con la extensión .txt
            ' Dim ficLocal As String = Path.Combine(Ruta, Path.GetFileName(ficFTP) & ".txt")
            Dim ficLocal As String = Path.Combine(Ruta, "DGLexistencias.csv")
            Dim sw As New StreamWriter(ficLocal, False, Encoding.Default)
            sw.Write(res)
            sw.Close()
            reader.Close()


        Catch ex As Exception
            Dim ola As String = ex.Message
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ''  fichero = "productE.xml"
        Me.Panel1.Visible = True
        Me.Button1.Enabled = False
        If work.IsBusy Then

        Else
            work.RunWorkerAsync()
        End If


    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Close()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        Me.Button1.Enabled = True
        work.CancelAsync()
        Me.Panel1.Visible = False

    End Sub
End Class