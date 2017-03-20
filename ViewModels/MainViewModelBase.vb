Public MustInherit Class MainViewModelBase(Of VM As New)
    Inherits ViewModelBase

    Private Shared _Instance As VM

    Public Shared ReadOnly Property Instance As VM
        Get
            If _Instance Is Nothing Then _Instance = New VM
            Return _Instance
        End Get
    End Property

    Protected Sub New()
        MyBase.New
        If _Instance IsNot Nothing Then Throw New InvalidOperationException("The MainViewModelBase class can only be inherited once per application.")
    End Sub

End Class
