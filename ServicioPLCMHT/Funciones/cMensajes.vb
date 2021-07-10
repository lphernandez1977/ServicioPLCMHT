Public Class cMensajes
    Private _Mensaje_Error As String
    Private _Mensaje_Alerta As String

    Public Property MensajesError() As String
        Get
            Return _Mensaje_Error
        End Get
        Set(value As String)
            _Mensaje_Error = value
        End Set
    End Property

    Public Property MensajesAlerta() As String
        Get
            Return _Mensaje_Alerta
        End Get
        Set(value As String)
            _Mensaje_Alerta = value
        End Set
    End Property

    Public Shared MensajeError As String = String.Empty

End Class
