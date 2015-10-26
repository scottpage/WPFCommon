Public Class MouseUtils
    Inherits DependencyObject



    Private Shared _MouseDownPosition As Point
    Private Shared _CurrentDragElement As FrameworkElement
    Private Shared _LastDragElement As FrameworkElement
    Private Shared ReadOnly _Dragables As New List(Of FrameworkElement)
    Private Shared ReadOnly _OnMouseDownRegistrars As New List(Of FrameworkElement)
    Private Shared ReadOnly _OnMouseUpRegistrars As New List(Of FrameworkElement)
    Private Shared ReadOnly _OnMouseMoveRegistrars As New List(Of FrameworkElement)

#Region "Drag"

    Public Shared ReadOnly DragProperty As DependencyProperty = DependencyProperty.RegisterAttached("Drag",
                                                        GetType(Boolean),
                                                        GetType(MouseUtils),
                                                        New FrameworkPropertyMetadata(False,
                                                        New PropertyChangedCallback(AddressOf MouseUtils.OnDragPropertyChanged)))

    Public Shared Function GetDrag(ByVal element As DependencyObject) As Boolean
        If Not TypeOf element Is FrameworkElement OrElse element Is Nothing Then Throw New ArgumentException("Drag cannot be used with an element of this type.", "element")
        Return DirectCast(element.GetValue(DragProperty), Boolean)
    End Function

    Public Shared Sub SetDrag(ByVal element As DependencyObject, ByVal value As Boolean)
        If Not TypeOf element Is FrameworkElement OrElse element Is Nothing Then Throw New ArgumentException("Drag cannot be used with an element of this type.", "element")
        element.SetValue(DragProperty, value)
    End Sub

    Public Shared Sub OnDragPropertyChanged(sender As Object, e As DependencyPropertyChangedEventArgs)
        If Not TypeOf sender Is FrameworkElement OrElse sender Is Nothing Then Return
        Dim fe = DirectCast(sender, FrameworkElement)
        Dim CanDrag = GetDrag(fe)
        If CanDrag Then
            RegisterForMouseEvents(fe)
        Else
            UnRegisterForMouseEvents(fe)
        End If
    End Sub

#End Region

    Private Shared Sub RegisterForMouseEvents(fe As FrameworkElement)
        Panel.SetZIndex(fe, Integer.MaxValue)
        ReorderZIndex()
        _Dragables.Add(fe)
        If Not _OnMouseDownRegistrars.Contains(fe) Then
            AddHandler fe.MouseLeftButtonDown, AddressOf Sender_OnLeftMouseButtonDown
            _OnMouseDownRegistrars.Add(fe)
        End If
        If Not _OnMouseUpRegistrars.Contains(fe) Then
            AddHandler fe.MouseLeftButtonUp, AddressOf Sender_OnLeftMouseButtonUp
            _OnMouseUpRegistrars.Add(fe)
        End If
        If Not _OnMouseMoveRegistrars.Contains(fe) Then
            AddHandler fe.MouseMove, AddressOf Sender_OnMouseMove
            _OnMouseMoveRegistrars.Add(fe)
        End If
    End Sub

    Private Shared Sub UnRegisterForMouseEvents(fe As FrameworkElement)
        _Dragables.Remove(fe)
        If _OnMouseDownRegistrars.Contains(fe) Then
            RemoveHandler fe.MouseLeftButtonDown, AddressOf Sender_OnLeftMouseButtonDown
            _OnMouseDownRegistrars.Remove(fe)
        End If
        If _OnMouseUpRegistrars.Contains(fe) Then
            RemoveHandler fe.MouseLeftButtonUp, AddressOf Sender_OnLeftMouseButtonUp
            _OnMouseUpRegistrars.Remove(fe)
        End If
        If _OnMouseMoveRegistrars.Contains(fe) Then
            RemoveHandler fe.MouseMove, AddressOf Sender_OnMouseMove
            _OnMouseMoveRegistrars.Remove(fe)
        End If
    End Sub

    Private Shared Sub ReorderZIndex()
        _Dragables.ForEach(Sub(fe2)
                               If fe2 IsNot _CurrentDragElement Then
                                   Dim fe2ZI = Panel.GetZIndex(fe2)
                                   If fe2ZI > Integer.MinValue Then
                                       Panel.SetZIndex(fe2, fe2ZI - 1)
                                   Else
                                       Panel.SetZIndex(fe2, Integer.MinValue)
                                   End If
                               End If
                           End Sub)
    End Sub

    Private Shared Function GetClosestElement(fe As FrameworkElement) As CloseElement
        Dim Closest As New List(Of FrameworkElement)
        Dim FEOffset = VisualTreeHelper.GetOffset(fe)
        Dim DeltaV As Vector = Nothing
        For Each el In _Dragables
            If fe Is el Then Continue For
            Dim elOffset = VisualTreeHelper.GetOffset(el)
            DeltaV = FEOffset - elOffset
            If DeltaV.X <= 10 Then
                Closest.Add(el)
            ElseIf DeltaV.Y <= 10 Then
                Closest.Add(el)
            End If
        Next
        Return New CloseElement(fe, Closest.FirstOrDefault, DeltaV)
    End Function

    Private Class CloseElement

        Public Sub New(s As FrameworkElement, c As FrameworkElement, d As Vector)
            Source = s
            Closest = c
            Delta = d
        End Sub

        Public Source As FrameworkElement
        Public Closest As FrameworkElement
        Public Delta As Vector

    End Class

    Private Shared Sub Sender_OnLeftMouseButtonDown(sender As Object, e As Input.MouseButtonEventArgs)
        If e.ClickCount > 1 Then Return
        Dim fe = DirectCast(sender, FrameworkElement)
        Dim ShouldDrag = GetDrag(fe)
        If ShouldDrag Then
            e.Handled = True
            _CurrentDragElement = fe
            fe.CaptureMouse()
            Panel.SetZIndex(fe, Integer.MaxValue)
            ReorderZIndex()
            _MouseDownPosition = e.GetPosition(fe)
        End If
    End Sub

    Private Shared Sub Sender_OnLeftMouseButtonUp(sender As Object, e As Input.MouseButtonEventArgs)
        Dim fe = DirectCast(sender, FrameworkElement)
        Dim ShouldDrag = GetDrag(fe) AndAlso fe.IsMouseCaptured
        If ShouldDrag Then
            e.Handled = True
            fe.ReleaseMouseCapture()
            Dim ClosestElement = GetClosestElement(fe)
            If ClosestElement IsNot Nothing Then
                Canvas.SetRight(fe, ClosestElement.Delta.X + fe.ActualWidth)
                Canvas.SetBottom(fe, ClosestElement.Delta.Y + fe.ActualHeight)
                Canvas.SetLeft(fe, ClosestElement.Delta.X)
                Canvas.SetTop(fe, ClosestElement.Delta.Y)
            End If
            _CurrentDragElement = Nothing
            _LastDragElement = fe
        End If
    End Sub

    Private Shared Sub Sender_OnMouseMove(sender As Object, e As Input.MouseEventArgs)
        Dim fe = DirectCast(sender, FrameworkElement)
        Dim ShouldDrag = GetDrag(fe) AndAlso fe.IsMouseCaptured
        If ShouldDrag Then
            e.Handled = True
            Dim CurMousePos = e.GetPosition(fe)
            Dim Delta = _MouseDownPosition - CurMousePos
            fe.Margin = New Thickness(fe.Margin.Left - Delta.X,
                                     fe.Margin.Top - Delta.Y,
                                     fe.Margin.Right + Delta.X,
                                     fe.Margin.Bottom + Delta.Y)
        Else

        End If
    End Sub

End Class
