Public Module ValueTypeExtensions

#Region "ToInt32"

    <Extension>
    Public Function ToInt32(value() As Byte, Optional index As Integer = 0, Optional reverse As Boolean = False) As Integer
        Dim Bytes(4) As Byte
        Array.Copy(value, index, Bytes, 0, 4)
        If reverse Then Bytes = Bytes.Reverse.ToArray
        Return BitConverter.ToInt32(Bytes, 0)
    End Function

    <Extension>
    Public Function ToInt32(value As Char) As Integer
        Return Convert.ToInt32(value)
    End Function

    <Extension>
    Public Function ToInt32(value As String) As Integer
        Return Convert.ToInt32(value)
    End Function

    <Extension>
    Public Function ToInt32(value As Decimal) As Integer
        Return Convert.ToInt32(value)
    End Function

    <Extension>
    Public Function ToInt32(value As Single) As Integer
        Return Convert.ToInt32(value)
    End Function

    <Extension>
    Public Function ToInt32(value As Double) As Integer
        Return Convert.ToInt32(value)
    End Function

#End Region

#Region "ToUInt32"

    <Extension>
    Public Function ToUInt32(value() As Byte, Optional index As Integer = 0, Optional reverse As Boolean = False) As UInteger
        Dim Bytes(4) As Byte
        Array.Copy(value, index, Bytes, 0, 4)
        If reverse Then Bytes = Bytes.Reverse.ToArray
        Return BitConverter.ToUInt32(Bytes, 0)
    End Function

    <Extension>
    Public Function ToUInt32(value As Char) As UInteger
        Return Convert.ToUInt32(value)
    End Function

    <Extension>
    Public Function ToUInt32(value As String) As UInteger
        Return Convert.ToUInt32(value)
    End Function

    <Extension>
    Public Function ToUInt32(value As Decimal) As UInteger
        Return Convert.ToUInt32(value)
    End Function

    <Extension>
    Public Function ToUInt32(value As Single) As UInteger
        Return Convert.ToUInt32(value)
    End Function

    <Extension>
    Public Function ToUInt32(value As Double) As UInteger
        Return Convert.ToUInt32(value)
    End Function

#End Region

End Module
