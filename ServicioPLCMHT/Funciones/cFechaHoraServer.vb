Public Class cFechaHoraServer
    Private _FechaLectura As DateTime

    Public Property FechaLectura() As DateTime
        Get
            Return _FechaLectura
        End Get
        Set(value As DateTime)
            _FechaLectura = value
        End Set
    End Property
End Class
