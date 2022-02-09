Public Class LineBreak
    Inherits CustomGesture

    Public Sub New()
        MyClass.New(Nothing)
    End Sub

    Public Sub New(ByVal strokeInfo As StrokeInfo)
        MyBase.New(strokeInfo)
        Name = "Linebreak"
    End Sub

    Protected Overrides Function Recognize() As Boolean
        Return _
            StrokeInfo.IsMatch(Vectors.StartTick & Vectors.Rights & _
                Vectors.Ups & Vectors.Rights & Vectors.EndTick) _
            AndAlso StrokeInfo.StrokeStatistics.Square >= 0.9 _
            AndAlso StrokeInfo.StrokeStatistics.RightUp < 0.1

    End Function

End Class
