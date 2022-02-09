Public Class RightBracket
    Inherits CustomGesture

    Public Sub New()
        MyClass.New(Nothing)
    End Sub

    Public Sub New(ByVal strokeInfo As StrokeInfo)
        MyBase.New(strokeInfo)
        Name = "Right Bracket"
    End Sub

    Protected Overrides Function Recognize() As Boolean
        Return _
            StrokeInfo.StrokeStatistics.StartEndProximity > CLOSED_PROXIMITY _
            AndAlso StrokeInfo.StrokeStatistics.StopPoints = 4 _
            AndAlso StrokeInfo.StrokeStatistics.Square > 0.9 _
            AndAlso StrokeInfo.IsMatch(Vectors.StartTick & Vectors.Rights & _
                Vectors.Downs & Vectors.Lefts & Vectors.EndTick, _
                0, False, True)
    End Function


End Class
