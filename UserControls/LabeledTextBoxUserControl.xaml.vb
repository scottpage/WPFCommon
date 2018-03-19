Public Class LabeledTextBoxUserControl

    Public Shared ReadOnly LabelTextProperty As DependencyProperty =
                           DependencyProperty.Register("LabelText",
                           GetType(String), GetType(LabeledTextBoxUserControl),
                           New PropertyMetadata(String.Empty))

    Public Property LabelText As String
        Get
            Return GetValue(LabelTextProperty).ToString
        End Get

        Set(ByVal value As String)
            SetValue(LabelTextProperty, value)
        End Set
    End Property

    Public Shared ReadOnly TextBoxTextProperty As DependencyProperty =
                           DependencyProperty.Register("TextBoxText",
                           GetType(String), GetType(LabeledTextBoxUserControl),
                           New PropertyMetadata(String.Empty))

    Public Property TextBoxText As String
        Get
            Return GetValue(TextBoxTextProperty).ToString
        End Get

        Set(ByVal value As String)
            SetValue(TextBoxTextProperty, value)
        End Set
    End Property

End Class
