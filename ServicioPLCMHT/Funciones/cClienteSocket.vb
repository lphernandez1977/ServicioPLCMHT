Imports System.Net.Sockets
Imports System.Threading
Imports System.Text
Imports System.IO

Public Class cClienteSocket
    Private mensajesEnviarRecibir As Stream 'Para enviar y recibir datos del servidor
    Private ipServidor As String 'Dirección IP
    Private puertoServidor As String 'Puerto de escucha

    Private clienteTCP As TcpClient
    Private hiloMensajeServidor As Thread 'Escuchar mensajes enviados desde el servidor

    Public Event ConexionTerminada()
    Public Event DatosRecibidos(ByVal datos As String)

    'Dirección IP del servidor al que nos conectaremos
    Public Property IP() As String
        Get
            IP = ipServidor
        End Get

        Set(ByVal Value As String)
            ipServidor = Value
        End Set
    End Property

    'Puerto por el que realizar la conexión al servidor
    Public Property Puerto() As Integer
        Get
            Puerto = puertoServidor
        End Get

        Set(ByVal Value As Integer)
            puertoServidor = Value
        End Set
    End Property

    'Procedimiento para realizar la conexión con el servidor
    Public Sub Conectar()
        clienteTCP = New TcpClient()

        'Conectar con el servidor
        clienteTCP.Connect(IP, Puerto)
        mensajesEnviarRecibir = clienteTCP.GetStream()

        'Crear hilo para establecer escucha de posibles mensajes
        'enviados por el servidor al cliente
        hiloMensajeServidor = New Thread(AddressOf LeerSocket)
        hiloMensajeServidor.Start()
    End Sub

    'Procedimiento para cerrar la conexión con el servidor
    Public Sub Desconectar()
        'desconectamos del servidor
        clienteTCP.Close()
        'abortamos el hilo (thread)
        hiloMensajeServidor.Abort()
    End Sub

    'Enviar mensaje al servidor
    Public Sub EnviarDatos(ByVal Datos As String)
        Dim BufferDeEscritura() As Byte

        BufferDeEscritura = Encoding.ASCII.GetBytes(Datos)
        If Not (mensajesEnviarRecibir Is Nothing) Then
            mensajesEnviarRecibir.Write(BufferDeEscritura,
                    0, BufferDeEscritura.Length)
        End If
    End Sub

    Private Sub LeerSocket()
        Dim BufferDeLectura() As Byte

        While True
            Try
                BufferDeLectura = New Byte(100) {}

                'Esperar a que llegue algún mensaje
                mensajesEnviarRecibir.Read(BufferDeLectura,
                      0, BufferDeLectura.Length)

                'Generar evento DatosRecibidos cuando se recibien datos desde el servidor
                RaiseEvent DatosRecibidos(Encoding.ASCII.GetString(BufferDeLectura))
            Catch e As Exception
                Exit While
            End Try
        End While

        'Finalizar conexión y generar evento ConexionTerminada
        RaiseEvent ConexionTerminada()
    End Sub


End Class
