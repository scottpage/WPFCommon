Imports MvvmValidation
Imports System.Collections.ObjectModel
Imports System.Reflection
Imports System.Threading.Tasks

''' <summary>
''' Base View Model class for implementing MVVM in WPF.
''' </summary>
''' <remarks></remarks>
Public Class ViewModelBase
    Implements INotifyPropertyChanging
    Implements INotifyPropertyChanged
    Implements IDisposable

    Protected Sub New()
    End Sub

    <XmlIgnore>
    Public ReadOnly Property ApplicationTitle As String
        Get
            Return My.Application.Info.Title
        End Get
    End Property

    <XmlIgnore>
    Private Shared _CreatorDispatcher As Dispatcher = Nothing
    Public Shared Property CreatorDispatcher As Dispatcher
        Get
            Return _CreatorDispatcher
        End Get
        Set(value As Dispatcher)
            If _CreatorDispatcher IsNot Nothing Then
                Throw New InvalidOperationException("The CreatorDispatcher can only be set once per application instance.")
            End If
            _CreatorDispatcher = value
        End Set
    End Property

    <XmlIgnore>
    Public Property IsEventsEnabled As Boolean = True

    Private _IsInitializing As Boolean
    Public ReadOnly Property IsInitializing As Boolean
        Get
            Return _IsInitializing
        End Get
    End Property

    Private _IsInitialized As Boolean
    Public ReadOnly Property IsInitialized As Boolean
        Get
            Return _IsInitialized
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

    Public Async Function Initialize(Of T)(options As T) As Task
        If CreatorDispatcher Is Nothing Then Throw New InvalidOperationException("ViewModelBase.CreatorDispatcher must be set before initializing.")
        SetProperty(Function() IsInitializing, _IsInitializing, True)
        IsEventsEnabled = False
        Await OnInitialize(options)
        Await OnInitialized()
        IsDirty = False
        IsEventsEnabled = True
        SetProperty(Function() IsInitializing, _IsInitializing, False)
        SetProperty(Function() IsInitialized, _IsInitialized, True)
    End Function

    Protected Overridable Async Function OnInitialize(Of T)(options As T) As Task
        Await Task.CompletedTask
    End Function

    Protected Overridable Async Function OnInitialized() As Task
        Await Task.CompletedTask
    End Function

    Public Function CheckAccess() As Boolean
        Return CreatorDispatcher.CheckAccess
    End Function

#Region "INotifyPropertyChanging"

    Public Event PropertyChanging(sender As Object, e As PropertyChangingEventArgs) Implements INotifyPropertyChanging.PropertyChanging

    Protected Overridable Sub OnPropertyChanging(Of T)(expression As Expression(Of Func(Of T)))
        Dim PropertyName = GetPropertyName(expression)
        OnPropertyChanging(PropertyName)
    End Sub

    Private Sub OnPropertyChanging(propertyName As String)
        If IsEventsEnabled Then OnPropertyChanging(Me, New PropertyChangingEventArgs(propertyName))
    End Sub

    Protected Overridable Sub OnPropertyChanging(Of T)(sender As Object, expression As Expression(Of Func(Of T)), oldValue As T, newValue As T)
        OnPropertyChanging(Me, New PropertyChangingExtendedEventArgs(Of T)(GetPropertyName(expression), oldValue, newValue))
    End Sub

    Public Overridable Sub OnPropertyChanging(sender As Object, e As PropertyChangingEventArgs)
        If CheckAccess() Then
            ValidatePropertyName(e.PropertyName)
            RaiseEvent PropertyChanging(Me, e)
        Else
            CreatorDispatcher.BeginInvoke(New Action(Of Object, PropertyChangingEventArgs)(AddressOf OnPropertyChanging), sender, e)
        End If
    End Sub

#End Region

#Region "INotifyPropertyChanged"

    Public Event PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Implements INotifyPropertyChanged.PropertyChanged

    Protected Overridable Sub OnPropertyChanged(Of T)(expression As Expression(Of Func(Of T)))
        Dim PropertyName = GetPropertyName(expression)
        OnPropertyChanged(PropertyName)
    End Sub

    Private Sub OnPropertyChanged(propertyName As String)
        If IsEventsEnabled Then OnPropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub

    Protected Overridable Sub OnPropertyChanged(Of T)(sender As Object, expression As Expression(Of Func(Of T)), oldValue As T, newValue As T)
        OnPropertyChanged(Me, New PropertyChangedExtendedEventArgs(Of T)(GetPropertyName(expression), oldValue, newValue))
    End Sub

    Public Overridable Sub OnPropertyChanged(sender As Object, e As PropertyChangedEventArgs)
        If CheckAccess() Then
            ValidatePropertyName(e.PropertyName)
            RaiseEvent PropertyChanged(Me, e)
        Else
            CreatorDispatcher.BeginInvoke(New Action(Of Object, PropertyChangedEventArgs)(AddressOf OnPropertyChanged), sender, e)
        End If
    End Sub

#End Region

#Region "SetProperty"

    Protected Sub SetProperty(Of T)(expression As Expression(Of Func(Of T)), ByRef backingField As T, newValue As T)
        If (backingField Is Nothing And newValue Is Nothing) OrElse (backingField IsNot Nothing AndAlso backingField.Equals(newValue)) Then Return
        Dim OldBackingFieldValue = backingField
        If IsEventsEnabled Then OnPropertyChanging(Me, expression, OldBackingFieldValue, newValue)
        backingField = newValue
        If IsEventsEnabled Then OnPropertyChanged(Me, expression, OldBackingFieldValue, newValue)
    End Sub

    Protected Function GetPropertyName(Of T)(expression As Expression(Of Func(Of T))) As String
        If expression Is Nothing Then Throw New ArgumentNullException("expression")
        Dim body = expression.Body
        Dim memberExpression As MemberExpression = TryCast(body, MemberExpression)
        If memberExpression Is Nothing Then memberExpression = DirectCast(DirectCast(body, UnaryExpression).Operand, MemberExpression)
        Return memberExpression.Member.Name
    End Function

#End Region

    Private Sub ValidatePropertyName(propertyName As String)
#If Not DEBUG Then
        Return  
#End If
        Dim RequestedPropInfo = Me.GetType.GetProperty(propertyName, BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.FlattenHierarchy Or BindingFlags.GetProperty)
        If RequestedPropInfo IsNot Nothing Then Return
        Throw New ArgumentException(String.Format("No property named, {0}, exists for the current object, {1}", propertyName, Me.GetType.FullName), propertyName)
    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
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
