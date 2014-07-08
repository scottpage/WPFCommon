Imports System.ComponentModel

''' <summary>
''' A base class for implementing a MVVM style in WPF.
''' </summary>
''' <remarks></remarks>
Public Class ViewModelBase
    Implements INotifyPropertyChanging
    Implements INotifyPropertyChanged

#Region "Constructors"

    Protected Sub New()
        _CreatorDispatcher = System.Windows.Threading.Dispatcher.CurrentDispatcher
    End Sub

#End Region

#Region "Properties"

    Private _CreatorDispatcher As Dispatcher
    <XmlIgnore>
    Protected ReadOnly Property CreatorDispatcher As Dispatcher
        Get
            Return _CreatorDispatcher
        End Get
    End Property

    Private _IsDirty As Boolean
    <XmlIgnore>
    Public Property IsDirty As Boolean
        Get
            Return _IsDirty
        End Get
        Set(ByVal Value As Boolean)
            SetProperty(Function() IsDirty, _IsDirty, Value)
        End Set
    End Property

    Private Shared _EventsEnabled As Boolean = True
    <XmlIgnore>
    Public Shared Property EventsEnabled As Boolean
        Get
            Return _EventsEnabled
        End Get
        Set(ByVal Value As Boolean)
            _EventsEnabled = Value
        End Set
    End Property

#End Region

#Region "Methods"

    Private Delegate Sub RefreshCommandsSyncHandler()

    Public Overridable Sub Initialize()
        OnInitialize()
    End Sub

    Protected Overridable Sub OnInitialize()
    End Sub

    Public Sub RefreshCommands()
        Dim Handler As New RefreshCommandsSyncHandler(AddressOf RefreshCommandsSync)
        CreatorDispatcher.Invoke(Handler)
    End Sub

    Protected Overridable Sub RefreshCommandsSync()
    End Sub

    Public Overridable Sub Reset()
    End Sub

#End Region

#Region "INotifyPropertyChanging and INotifyPropertyChanged"

    Public Event PropertyChanging(sender As Object, e As System.ComponentModel.PropertyChangingEventArgs) Implements System.ComponentModel.INotifyPropertyChanging.PropertyChanging
    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    Private Delegate Sub PropertyChangeDelegate(propertyName As String)

    Protected Overridable Sub UpdateProperty(propertyName As String, ByRef backingFieldValue As Object, newValue As Object)
        If backingFieldValue IsNot Nothing AndAlso backingFieldValue.Equals(newValue) Then Return
        OnPropertyChanging(propertyName)
        backingFieldValue = newValue
        OnPropertyChanged(propertyName)
    End Sub

    Protected Overridable Sub UpdateProperty(propertyName As String, ByRef backingFieldValue As String, newValue As String)
        If Not String.IsNullOrEmpty(backingFieldValue) AndAlso backingFieldValue.Equals(newValue) Then Return
        OnPropertyChanging(propertyName)
        backingFieldValue = newValue
        OnPropertyChanged(propertyName)
    End Sub

    Protected Overridable Sub UpdateProperty(propertyName As String, ByRef backingFieldValue As Double, newValue As Double)
        If backingFieldValue.Equals(newValue) Then Return
        OnPropertyChanging(propertyName)
        backingFieldValue = newValue
        OnPropertyChanged(propertyName)
    End Sub

    Protected Overridable Sub UpdateProperty(propertyName As String, ByRef backingFieldValue As Single, newValue As Single)
        If backingFieldValue.Equals(newValue) Then Return
        OnPropertyChanging(propertyName)
        backingFieldValue = newValue
        OnPropertyChanged(propertyName)
    End Sub

    Protected Overridable Sub UpdateProperty(propertyName As String, ByRef backingFieldValue As Long, newValue As Long)
        If backingFieldValue.Equals(newValue) Then Return
        OnPropertyChanging(propertyName)
        backingFieldValue = newValue
        OnPropertyChanged(propertyName)
    End Sub

    Protected Overridable Sub UpdateProperty(propertyName As String, ByRef backingFieldValue As ULong, newValue As ULong)
        If backingFieldValue.Equals(newValue) Then Return
        OnPropertyChanging(propertyName)
        backingFieldValue = newValue
        OnPropertyChanged(propertyName)
    End Sub

    Protected Overridable Sub UpdateProperty(propertyName As String, ByRef backingFieldValue As Integer, newValue As Integer)
        If backingFieldValue.Equals(newValue) Then Return
        OnPropertyChanging(propertyName)
        backingFieldValue = newValue
        OnPropertyChanged(propertyName)
    End Sub

    Protected Overridable Sub UpdateProperty(propertyName As String, ByRef backingFieldValue As UInteger, newValue As UInteger)
        If backingFieldValue.Equals(newValue) Then Return
        OnPropertyChanging(propertyName)
        backingFieldValue = newValue
        OnPropertyChanged(propertyName)
    End Sub

    Protected Overridable Sub UpdateProperty(propertyName As String, ByRef backingFieldValue As Short, newValue As Short)
        If backingFieldValue.Equals(newValue) Then Return
        OnPropertyChanging(propertyName)
        backingFieldValue = newValue
        OnPropertyChanged(propertyName)
    End Sub

    Protected Overridable Sub UpdateProperty(propertyName As String, ByRef backingFieldValue As UShort, newValue As UShort)
        If backingFieldValue.Equals(newValue) Then Return
        OnPropertyChanging(propertyName)
        backingFieldValue = newValue
        OnPropertyChanged(propertyName)
    End Sub

    Protected Overridable Sub UpdateProperty(propertyName As String, ByRef backingFieldValue As Boolean, newValue As Boolean)
        If backingFieldValue.Equals(newValue) Then Return
        OnPropertyChanging(propertyName)
        backingFieldValue = newValue
        OnPropertyChanged(propertyName)
    End Sub

#Region "New Lamba way"

    Protected Sub SetProperty(Of T)(expression As Expression(Of Func(Of T)), ByRef field As T, value As T)
        If field IsNot Nothing AndAlso field.Equals(value) Then Return
        Dim oldValue = field
        OnPropertyChanging(Me, expression, oldValue, value)
        field = value
        OnPropertyChanged(Me, expression, oldValue, value)
    End Sub

    Protected Sub OnPropertyChanging(Of T)(expression As Expression(Of Func(Of T)))
        Dim PropertyName = GetPropertyName(expression)
        OnPropertyChanging(PropertyName)
    End Sub

    Protected Overridable Sub OnPropertyChanging(Of T)(sender As Object, expression As Expression(Of Func(Of T)), oldValue As T, newValue As T)
        OnPropertyChanging(Me, New PropertyChangingExtendedEventArgs(Of T)(GetPropertyName(expression), oldValue, newValue))
    End Sub

    Protected Overridable Sub OnPropertyChanging(sender As Object, e As PropertyChangingEventArgs)
        If System.Windows.Threading.Dispatcher.CurrentDispatcher Is CreatorDispatcher Then
            If EventsEnabled Then RaiseEvent PropertyChanging(Me, e)
        Else
            Dim Handler As New PropertyChangingEventHandler(AddressOf OnPropertyChanging)
            CreatorDispatcher.BeginInvoke(Handler, sender, e)
        End If
    End Sub

    Protected Sub OnPropertyChanged(Of T)(expression As Expression(Of Func(Of T)))
        Dim PropertyName = GetPropertyName(expression)
        OnPropertyChanged(PropertyName)
    End Sub

    Protected Overridable Sub OnPropertyChanged(Of T)(sender As Object, expression As Expression(Of Func(Of T)), oldValue As T, newValue As T)
        OnPropertyChanged(Me, New PropertyChangedExtendedEventArgs(Of T)(GetPropertyName(expression), oldValue, newValue))
    End Sub

    Public Overridable Sub OnPropertyChanged(sender As Object, e As PropertyChangedEventArgs)
        If System.Windows.Threading.Dispatcher.CurrentDispatcher Is CreatorDispatcher Then
            If EventsEnabled Then RaiseEvent PropertyChanged(Me, e)
        Else
            Dim Handler As New PropertyChangedEventHandler(AddressOf OnPropertyChanged)
            CreatorDispatcher.BeginInvoke(Handler, sender, e)
        End If
    End Sub

    Protected Function GetPropertyName(Of T)(expression As Expression(Of Func(Of T))) As String
        Dim memberExp = DirectCast(expression.Body, MemberExpression)
        Return memberExp.Member.Name
    End Function

    Protected Sub OnPropertyChangingAndChanged(Of T)(expression As Expression(Of Func(Of T)))
        Dim PropertyName = GetPropertyName(expression)
        OnPropertyChanging(PropertyName)
        OnPropertyChanged(PropertyName)
    End Sub

#End Region

    Protected Overridable Sub OnPropertyChanging(propertyName As String)
#If DEBUG Then
#If VALIDATEPROPERTYNAMES Then
        ValidatePropertyName(propertyName)
#End If
#End If
        If System.Windows.Threading.Dispatcher.CurrentDispatcher Is CreatorDispatcher Then
            If EventsEnabled Then RaiseEvent PropertyChanging(Me, New PropertyChangingEventArgs(propertyName))
        Else
            Dim Handler As New PropertyChangeDelegate(AddressOf OnPropertyChanging)
            CreatorDispatcher.BeginInvoke(Handler, propertyName)
        End If
    End Sub

    Protected Overridable Sub OnPropertyChanged(propertyName As String)
#If DEBUG Then
#If VALIDATEPROPERTYNAMES Then
        ValidatePropertyName(propertyName)
#End If
#End If
        If System.Windows.Threading.Dispatcher.CurrentDispatcher Is CreatorDispatcher Then
            If EventsEnabled Then RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
        Else
            Dim Handler As New PropertyChangeDelegate(AddressOf OnPropertyChanged)
            CreatorDispatcher.BeginInvoke(Handler, propertyName)
        End If

        IsDirty = True
    End Sub

#If DEBUG Then

    Private Sub ValidatePropertyName(propertyName As String)
        If TypeDescriptor.GetProperties(Me).Find(propertyName, False) Is Nothing Then Throw New ArgumentException(String.Format("No property named ""{0}"" exists for the current object", propertyName), propertyName)
    End Sub

#End If

#End Region

End Class
