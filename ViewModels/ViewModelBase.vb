Imports MvvmValidation
Imports System.Collections.ObjectModel

''' <summary>
''' Base View Model class for implementing MVVM in WPF.
''' </summary>
''' <remarks></remarks>
Public Class ViewModelBase
    Implements INotifyPropertyChanging
    Implements INotifyPropertyChanged
    Implements INotifyDataErrorInfo
    Implements IDisposable

    Protected ReadOnly Property Validator As ValidationHelper
    Private WithEvents _NotifyDataErrorInfoAdapter As NotifyDataErrorInfoAdapter

    Private Shared ReadOnly _ViewModels As New ObservableCollection(Of ViewModelBase)
    Private Shared _AllViewModelsAreShutdown As Boolean = False

    Public ReadOnly Property ApplicationTitle As String
        Get
            Return My.Application.Info.Title
        End Get
    End Property

    Public Shared ReadOnly Property ViewModels As IEnumerable(Of ViewModelBase)
        Get
            Return _ViewModels.AsEnumerable
        End Get
    End Property

    Private Shared Sub AddViewModel(vm As ViewModelBase)
        If _AllViewModelsAreShutdown Then
            Throw New InvalidOperationException("The view models are shutdown and can no longer be used.")
            Return
        End If
        SyncLock _ViewModelsLockObj
            _ViewModels.Add(vm)
        End SyncLock
    End Sub

    Private Shared Sub RemoveViewModel(vm As ViewModelBase)
        SyncLock _ViewModelsLockObj
            _ViewModels.Remove(vm)
        End SyncLock
    End Sub

    Private Shared ReadOnly _ViewModelsLockObj As New Object
    Public Shared Sub ShutdownViewModels()
        SyncLock _ViewModelsLockObj
            If _AllViewModelsAreShutdown Then Return
            For Each VM In _ViewModels.ToList
                VM.Shutdown()
                _ViewModels.Remove(VM)
            Next
            _ViewModels.Clear()
            _AllViewModelsAreShutdown = True
        End SyncLock
    End Sub

#Region "Constructors"

    Protected Sub New()
        CreatorDispatcher = Dispatcher.CurrentDispatcher
        AddViewModel(Me)
    End Sub

    <XmlIgnore>
    Protected ReadOnly Property CreatorDispatcher As Dispatcher

    Private _IsShutdown As Boolean
    Public ReadOnly Property IsShutdown As Boolean
        Get
            Return _IsShutdown
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

    <XmlIgnore>
    Public Shared Property EventsEnabled As Boolean = True

#End Region

#Region "Methods"

    Private Delegate Sub RefreshCommandsSyncHandler()

    Public Sub Initialize()
        OnInitialize()
    End Sub

    Protected Overridable Sub OnInitialize()
        OnInitialized()
    End Sub

    Protected Overridable Sub OnInitialized()
    End Sub

    Public Sub RefreshCommands()
        Dim Handler As New RefreshCommandsSyncHandler(AddressOf RefreshCommandsSync)
        CreatorDispatcher.Invoke(Handler)
    End Sub

    Protected Overridable Sub RefreshCommandsSync()
    End Sub

    Public Overridable Sub Reset()
    End Sub

    Public Sub Shutdown()
        OnShutdown()
        SetProperty(Function() IsShutdown, _IsShutdown, True)
    End Sub

    Protected Overridable Sub OnShutdown()
    End Sub

    Private _Creating As Boolean = False
    ''' <summary>
    ''' Call BeginCreate when creating a new instance requires setting properties that call SetProperty (notify).  Remember to call EndCreate when completed.
    ''' </summary>
    Public Sub BeginCreate()
        _Creating = True
    End Sub

    Public Sub EndCreate()
        _Creating = False
    End Sub

#End Region

#Region "INotifyPropertyChanging and INotifyPropertyChanged"

    Public Event PropertyChanging(sender As Object, e As System.ComponentModel.PropertyChangingEventArgs) Implements System.ComponentModel.INotifyPropertyChanging.PropertyChanging
    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    Protected Sub NotifyPropertyChanged(propertyName As String)
        If Not EventsEnabled Or _Creating Then Return
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub

    '#Region "Obsolete UpdateProperty methods"

    '    <Obsolete("Use SetProperty")>
    '    Protected Overridable Sub UpdateProperty(propertyName As String, ByRef backingFieldValue As Object, newValue As Object)
    '        If backingFieldValue IsNot Nothing AndAlso backingFieldValue.Equals(newValue) Then Return
    '        OnPropertyChanging(propertyName)
    '        backingFieldValue = newValue
    '        OnPropertyChanged(propertyName)
    '    End Sub

    '    <Obsolete("Use SetProperty")>
    '    Protected Overridable Sub UpdateProperty(propertyName As String, ByRef backingFieldValue As String, newValue As String)
    '        If Not String.IsNullOrEmpty(backingFieldValue) AndAlso backingFieldValue.Equals(newValue) Then Return
    '        OnPropertyChanging(propertyName)
    '        backingFieldValue = newValue
    '        OnPropertyChanged(propertyName)
    '    End Sub

    '    <Obsolete("Use SetProperty")>
    '    Protected Overridable Sub UpdateProperty(propertyName As String, ByRef backingFieldValue As Double, newValue As Double)
    '        If backingFieldValue.Equals(newValue) Then Return
    '        OnPropertyChanging(propertyName)
    '        backingFieldValue = newValue
    '        OnPropertyChanged(propertyName)
    '    End Sub

    '    <Obsolete("Use SetProperty")>
    '    Protected Overridable Sub UpdateProperty(propertyName As String, ByRef backingFieldValue As Single, newValue As Single)
    '        If backingFieldValue.Equals(newValue) Then Return
    '        OnPropertyChanging(propertyName)
    '        backingFieldValue = newValue
    '        OnPropertyChanged(propertyName)
    '    End Sub

    '    <Obsolete("Use SetProperty")>
    '    Protected Overridable Sub UpdateProperty(propertyName As String, ByRef backingFieldValue As Long, newValue As Long)
    '        If backingFieldValue.Equals(newValue) Then Return
    '        OnPropertyChanging(propertyName)
    '        backingFieldValue = newValue
    '        OnPropertyChanged(propertyName)
    '    End Sub

    '    <Obsolete("Use SetProperty")>
    '    Protected Overridable Sub UpdateProperty(propertyName As String, ByRef backingFieldValue As ULong, newValue As ULong)
    '        If backingFieldValue.Equals(newValue) Then Return
    '        OnPropertyChanging(propertyName)
    '        backingFieldValue = newValue
    '        OnPropertyChanged(propertyName)
    '    End Sub

    '    <Obsolete("Use SetProperty")>
    '    Protected Overridable Sub UpdateProperty(propertyName As String, ByRef backingFieldValue As Integer, newValue As Integer)
    '        If backingFieldValue.Equals(newValue) Then Return
    '        OnPropertyChanging(propertyName)
    '        backingFieldValue = newValue
    '        OnPropertyChanged(propertyName)
    '    End Sub

    '    <Obsolete("Use SetProperty")>
    '    Protected Overridable Sub UpdateProperty(propertyName As String, ByRef backingFieldValue As UInteger, newValue As UInteger)
    '        If backingFieldValue.Equals(newValue) Then Return
    '        OnPropertyChanging(propertyName)
    '        backingFieldValue = newValue
    '        OnPropertyChanged(propertyName)
    '    End Sub

    '    <Obsolete("Use SetProperty")>
    '    Protected Overridable Sub UpdateProperty(propertyName As String, ByRef backingFieldValue As Short, newValue As Short)
    '        If backingFieldValue.Equals(newValue) Then Return
    '        OnPropertyChanging(propertyName)
    '        backingFieldValue = newValue
    '        OnPropertyChanged(propertyName)
    '    End Sub

    '    <Obsolete("Use SetProperty")>
    '    Protected Overridable Sub UpdateProperty(propertyName As String, ByRef backingFieldValue As UShort, newValue As UShort)
    '        If backingFieldValue.Equals(newValue) Then Return
    '        OnPropertyChanging(propertyName)
    '        backingFieldValue = newValue
    '        OnPropertyChanged(propertyName)
    '    End Sub

    '    <Obsolete("Use SetProperty")>
    '    Protected Overridable Sub UpdateProperty(propertyName As String, ByRef backingFieldValue As Boolean, newValue As Boolean)
    '        If backingFieldValue.Equals(newValue) Then Return
    '        OnPropertyChanging(propertyName)
    '        backingFieldValue = newValue
    '        OnPropertyChanged(propertyName)
    '    End Sub

    '#End Region

#Region "New Lamba SetProperty"

    Protected Sub SetProperty(Of T)(expression As Expression(Of Func(Of T)), ByRef field As T, value As T)
        If field IsNot Nothing AndAlso field.Equals(value) Then Return
        Dim oldValue = field
        If EventsEnabled And Not _Creating Then OnPropertyChanging(Me, expression, oldValue, value)
        field = value
        If EventsEnabled And Not _Creating Then OnPropertyChanged(Me, expression, oldValue, value)
    End Sub

    Protected Overridable Sub OnPropertyChanging(Of T)(expression As Expression(Of Func(Of T)))
        Dim PropertyName = GetPropertyName(expression)
        OnPropertyChanging(PropertyName)
    End Sub

    Protected Overridable Sub OnPropertyChanging(Of T)(sender As Object, expression As Expression(Of Func(Of T)), oldValue As T, newValue As T)
        OnPropertyChanging(Me, New PropertyChangingExtendedEventArgs(Of T)(GetPropertyName(expression), oldValue, newValue))
    End Sub

    Protected Overridable Sub OnPropertyChanging(sender As Object, e As PropertyChangingEventArgs)
        If System.Windows.Threading.Dispatcher.CurrentDispatcher Is CreatorDispatcher Then
            RaiseEvent PropertyChanging(Me, e)
        Else
            Dim Handler As New PropertyChangingEventHandler(AddressOf OnPropertyChanging)
            CreatorDispatcher.BeginInvoke(Handler, sender, e)
        End If
    End Sub

    Protected Overridable Sub OnPropertyChanged(Of T)(expression As Expression(Of Func(Of T)))
        Dim PropertyName = GetPropertyName(expression)
        OnPropertyChanged(PropertyName)
    End Sub

    Protected Overridable Sub OnPropertyChanged(Of T)(sender As Object, expression As Expression(Of Func(Of T)), oldValue As T, newValue As T)
        OnPropertyChanged(Me, New PropertyChangedExtendedEventArgs(Of T)(GetPropertyName(expression), oldValue, newValue))
    End Sub

    Public Overridable Sub OnPropertyChanged(sender As Object, e As PropertyChangedEventArgs)
        If System.Windows.Threading.Dispatcher.CurrentDispatcher Is CreatorDispatcher Then
            RaiseEvent PropertyChanged(Me, e)
        Else
            Dim Handler As New PropertyChangedEventHandler(AddressOf OnPropertyChanged)
            CreatorDispatcher.BeginInvoke(Handler, sender, e)
        End If
    End Sub

    Protected Function GetPropertyName(Of T)(expression As Expression(Of Func(Of T))) As String
        If expression Is Nothing Then
            Throw New ArgumentNullException("expression")
        End If

        Dim body = expression.Body
        Dim memberExpression As MemberExpression = TryCast(body, MemberExpression)
        If memberExpression Is Nothing Then
            memberExpression = DirectCast(DirectCast(body, UnaryExpression).Operand, MemberExpression)
        End If
        Return memberExpression.Member.Name
    End Function

    'Original version, see more robust revision above.
    'Protected Function GetPropertyName(Of T)(expression As Expression(Of Func(Of T))) As String
    '    Dim memberExp = DirectCast(expression.Body, MemberExpression)
    '    Return memberExp.Member.Name
    'End Function

    Protected Sub OnPropertyChangingAndChanged(Of T)(expression As Expression(Of Func(Of T)))
        Dim PropertyName = GetPropertyName(expression)
        OnPropertyChanging(PropertyName)
        OnPropertyChanged(PropertyName)
    End Sub

#End Region

    Private Delegate Sub PropertyChangeDelegate(propertyName As String)

    Private Sub OnPropertyChanging(propertyName As String)
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

    Private Sub OnPropertyChanged(propertyName As String)
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
        If TypeDescriptor.GetProperties(Me).Find(propertyName, False) Is Nothing Then
            Throw New ArgumentException(String.Format("No property named, {0}, exists for the current object, {1}", propertyName, Me.GetType.FullName), propertyName)
        End If
    End Sub

#End If

#End Region

#Region "INotifyDataErrorInfo and MvvmValidation"

    Public Event ErrorsChanged As EventHandler(Of DataErrorsChangedEventArgs) Implements INotifyDataErrorInfo.ErrorsChanged

    Private Sub _NotifyDataErrorInfoAdapter_ErrorsChanged(sender As Object, e As DataErrorsChangedEventArgs) Handles _NotifyDataErrorInfoAdapter.ErrorsChanged
        RaiseEvent ErrorsChanged(Me, e)
    End Sub

    Public ReadOnly Property HasErrors As Boolean Implements INotifyDataErrorInfo.HasErrors
        Get
            Return _NotifyDataErrorInfoAdapter.HasErrors
        End Get
    End Property

    Public Function GetErrors(propertyName As String) As IEnumerable Implements INotifyDataErrorInfo.GetErrors
        Return _NotifyDataErrorInfoAdapter.GetErrors(propertyName)
    End Function

#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                RemoveViewModel(Me)
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        ' GC.SuppressFinalize(Me)
    End Sub

#End Region

End Class
