Public Class cEtiquetas
    Private _Etiqueta As String
    Private _Destino As String
    Private _Contador As Integer
    Private _FechaLectura As DateTime
    Private _DtdAnclaje As String
    Private _MensajeError As String
    Private _Mensaje As String
    Private _Estado As String


    Public Property Etiqueta() As String
        Get
            Return _Etiqueta
        End Get
        Set(value As String)
            _Etiqueta = value
        End Set
    End Property

    Public Property Destino() As String
        Get
            Return _Destino
        End Get
        Set(value As String)
            _Destino = value
        End Set
    End Property

    Public Property Contador() As Integer
        Get
            Return _Contador
        End Get
        Set(value As Integer)
            _Contador = value
        End Set
    End Property

    Public Property FechaLectura() As DateTime
        Get
            Return _FechaLectura
        End Get
        Set(value As DateTime)
            _FechaLectura = value
        End Set
    End Property

    Public Property DtdAnclaje() As String
        Get
            Return _DtdAnclaje
        End Get
        Set(value As String)
            _DtdAnclaje = value
        End Set
    End Property

    Public Property MensajeError() As String
        Get
            Return _MensajeError
        End Get
        Set(value As String)
            _MensajeError = value
        End Set
    End Property

    Public Property Mensaje() As String
        Get
            Return _Mensaje
        End Get
        Set(value As String)
            _Mensaje = value
        End Set
    End Property

    Public Property Estado() As String
        Get
            Return _Estado
        End Get
        Set(value As String)
            _Estado = value
        End Set
    End Property



End Class
