Imports Microsoft.Ink

Public MustInherit Class CustomGesture
    <CLSCompliant(False)> Protected _name As String
    <CLSCompliant(False)> Protected _recognized As Boolean = False
    <CLSCompliant(False)> Protected _isMatch As Boolean = False
    <CLSCompliant(False)> Protected _strokeInfo As StrokeInfo

    Protected Const CLOSED_PROXIMITY As Double = 700

    Public Property StrokeInfo() As StrokeInfo
        Get
            Return _strokeInfo
        End Get
        Set(ByVal Value As StrokeInfo)
            _recognized = False
            _strokeInfo = Value
        End Set
    End Property

    Public Sub New(ByVal strokeInfo As StrokeInfo)
        Me.StrokeInfo = strokeInfo
    End Sub

    Public Overridable ReadOnly Property IsMatch() As Boolean
        Get
            If Not _recognized Then
                _isMatch = Recognize()
                _recognized = True
            End If
            Return _isMatch
        End Get
    End Property

    Public Property Name() As String
        Get
            Return _name
        End Get
        Set(ByVal Value As String)
            _name = Value
        End Set
    End Property

    Protected Overridable Function Recognize() As Boolean
        Return False
    End Function

End Class
