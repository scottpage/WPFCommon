Public Class PropertyChangingExtendedEventArgs(Of T)
    Inherits PropertyChangingEventArgs

    Public Sub New(propertyName As String, oldValue As T, newValue As T)
        MyBase.New(propertyName)
        _OldValue = oldValue
        _NewValue = newValue
    End Sub

    Private _OldValue As T
    Public ReadOnly Property OldValue() As T
        Get
            Return _OldValue
        End Get
    End Property

    Private _NewValue As T
    Public ReadOnly Property NewValue() As T
        Get
            Return _NewValue
        End Get
    End Property

End Class