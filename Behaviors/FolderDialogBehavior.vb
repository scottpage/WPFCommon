Imports System.Reflection

Public Class FolderDialogBehavior
    Inherits Behavior(Of Button)

    Public Property SetterName() As String
        Get
            Return m_SetterName
        End Get
        Set
            m_SetterName = Value
        End Set
    End Property
    Private m_SetterName As String

    Protected Overrides Sub OnAttached()
        MyBase.OnAttached()
        AddHandler AssociatedObject.Click, AddressOf OnClick
    End Sub

    Protected Overrides Sub OnDetaching()
        RemoveHandler AssociatedObject.Click, AddressOf OnClick
    End Sub

    Private Sub OnClick(sender As Object, e As RoutedEventArgs)
        Dim dialog = New Forms.FolderBrowserDialog()
        Dim result = dialog.ShowDialog()
        If result = Forms.DialogResult.OK AndAlso AssociatedObject.DataContext IsNot Nothing Then
            Dim propertyInfo = AssociatedObject.DataContext.[GetType]().GetProperties(BindingFlags.Instance Or BindingFlags.[Public]).Where(Function(p) p.CanRead AndAlso p.CanWrite).Where(Function(p) p.Name.Equals(SetterName)).First()

            propertyInfo.SetValue(AssociatedObject.DataContext, dialog.SelectedPath, Nothing)
        End If
    End Sub
End Class