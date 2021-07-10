Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Text
Imports System.IO
Imports System.Net.Sockets

Public Class ServicioPLCMHT

    Dim WithEvents socketCliente As New cClienteSocket
    Dim oMensajes As New cMensajes
    Dim tiempo As Integer = 0
    Dim oCxn As Boolean
    Dim oLog As New cRegistroLog

    'hilo timer
    Public Delegado As Threading.TimerCallback
    Public _Timer As System.Threading.Timer

    Protected Overrides Sub OnStart(ByVal args() As String)

        ' Agregue el código aquí para iniciar el servicio. Este método debería poner
        ' en movimiento los elementos para que el servicio pueda funcionar.
        If (Not System.Diagnostics.EventLog.SourceExists("ServicioPLCMHT")) Then
            System.Diagnostics.EventLog.CreateEventSource("EmacPLC", "Application")
        End If
        Registros.Source = "ServicioPLCMHT"
        Registros.Log = "Application"

        socketCliente.IP = ConfigurationManager.AppSettings("Ip")
        socketCliente.Puerto = ConfigurationManager.AppSettings("Puerto")
        tiempo = Convert.ToInt16(ConfigurationManager.AppSettings("tiempo"))

        oCxn = False

        ConectarTCP()

        Delegado = AddressOf Tarea
        _Timer = New System.Threading.Timer(Delegado, Nothing, tiempo, tiempo)
        _Timer.Change(tiempo, tiempo)

        Registros.WriteEntry("Servicio PLCMHT Iniciado en forma correcta")
        'oLog.RegistroLog("Servicio PLCMHT Iniciado en forma correcta")

    End Sub

    Protected Overrides Sub OnStop()
        ' Agregue el código aquí para realizar cualquier anulación necesaria para detener el servicio.
        DesconectarTCP()

        _Timer.Dispose()

        Registros.WriteEntry("Servicio PLCMHT Terminado en forma correcta")

    End Sub

    Private Sub ConectarTCP()

        Try
            socketCliente.Conectar()
            Registros.WriteEntry("Servicio conectado al EIS IP " + socketCliente.IP)
            oCxn = True
        Catch ex As Exception            
            Registros.WriteEntry(ex.Message.ToString())
        End Try

    End Sub

    Private Sub DesconectarTCP()
        Try
            socketCliente.Desconectar()
        Catch ex As Exception
            Registros.WriteEntry(ex.Message.ToString() + "Terminar TCPIP")
        End Try

    End Sub

    Private Sub Tarea(ByVal state As Object)

        _Timer.Change(-1, -1)

        If oCxn Then
            LecturaDB()
        Else
            ConectarTCP()
        End If

        _Timer.Change(tiempo, tiempo)

    End Sub

    Public Function IsPortOpen(ByVal Host As String, ByVal Port As Integer) As Boolean
        Dim m_sck As New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
        Try
            m_sck.Connect(Host, Port)
            Return True
        Catch ex As SocketException
            'Código para manejar error del socket (cerrado, conexión rechazada)
            ConectarTCP()
            Return True
        Catch ex As Exception
            'Código para manejar otra excepción
            oMensajes.MensajesError = ex.Message.ToString()
            Registros.WriteEntry(oMensajes.MensajesError)
        End Try
        Return False
    End Function

    Private Sub LecturaDB()
        Dim res As New cEtiquetas
        Dim upd As Integer
        Dim mensajeFinal As String

        Try
            If (oCxn) Then
                res = SeleccionEtiquetaAnclaje()
                mensajeFinal = res.Mensaje

                If ((res.Etiqueta <> Nothing)) Then
                    'enviar datos al EIS
                    socketCliente.EnviarDatos(mensajeFinal)

                    'actualiza etiqueta enviada
                    upd = Update_CartonEnviado(res.Contador)

                    If (upd >= 1) Then
                        Registros.WriteEntry(mensajeFinal)
                        'oLog.RegistroLog(mensajeFinal)
                    Else
                        Registros.WriteEntry(res.Etiqueta + " - " + cMensajes.MensajeError)
                    End If
                Else
                End If
            Else
            End If
        Catch ex As Exception
            oMensajes.MensajesError = ex.Message.ToString
            Registros.WriteEntry(oMensajes.MensajesError + " " + res.Etiqueta)
            'oLog.RegistroLog(oMensajes.MensajesError)
        End Try
    End Sub

#Region "BASE DATOS"

    Public Function SeleccionEtiquetaAnclaje() As cEtiquetas
        Dim cnx As SqlConnection
        Dim cmd As SqlCommand
        Dim dataReader As SqlDataReader
        Dim oEtiq As New cEtiquetas

        cnx = New SqlConnection(ConfigurationManager.ConnectionStrings("cnnString").ToString())
        cmd = New SqlCommand("sp_selec_CartonesLpn_Derivados", cnx)
        cmd.CommandType = CommandType.StoredProcedure
        Try
            cnx.Open()
            dataReader = cmd.ExecuteReader

            While dataReader.Read()                
                oEtiq.Contador = Convert.ToInt32(dataReader("Id"))
                oEtiq.Etiqueta = dataReader("CartonLPN")
                oEtiq.DtdAnclaje = dataReader("LabelCondition")
                oEtiq.Destino = dataReader("Store")
                oEtiq.Mensaje = dataReader("Mensaje")
            End While
            cnx.Close()
            cnx.Dispose()
            Return oEtiq

        Catch ex As Exception

            oEtiq.MensajeError = ex.Message.ToString()
            Return oEtiq

        End Try

    End Function

    Public Function Update_CartonEnviado(ByVal oCarton As Integer) As Integer
        Dim cnx As SqlConnection
        Dim cmd As SqlCommand
        Dim exec As Integer

        Try
            cnx = New SqlConnection(ConfigurationManager.ConnectionStrings("cnnString").ToString())
            cmd = New SqlCommand("sp_Modifica_CartonesLpn_Derivados", cnx)
            cmd.CommandType = CommandType.StoredProcedure
            'parametros de entrada
            cmd.Parameters.Add("@pContador", SqlDbType.Int).Value = oCarton
            cnx.Open()
            exec = cmd.ExecuteNonQuery
            cnx.Close()
            cnx.Dispose()
            Return exec
        Catch ex As Exception
            cMensajes.MensajeError = ex.Message.ToString()
            Return 0
        End Try
    End Function

#End Region

End Class
