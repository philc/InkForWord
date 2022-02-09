Public Class Lowercase
    Inherits CustomGesture

    Public Sub New()
        MyClass.New(Nothing)
    End Sub

    Public Sub New(ByVal strokeInfo As StrokeInfo)
        MyBase.New(strokeInfo)
        Name = "Lowercase"
    End Sub

    Protected Overrides Function Recognize() As Boolean
        Return _
            ( _
                StrokeInfo.IsMatch(Vectors.StartTick & Vectors.RightUps & Vectors.Rights & Vectors.EndTick) _
                OrElse StrokeInfo.IsMatch(Vectors.StartTick & Vectors.RightUps & Vectors.Rights & Vectors.EndTick) _
            ) _
            AndAlso StrokeInfo.StrokeStatistics.Square < 0.8 _
            AndAlso StrokeInfo.StrokeStatistics.Right > 0.25

    End Function

End Class
