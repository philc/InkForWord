Imports Microsoft.Ink
Imports System.Text.RegularExpressions

Public Class Circle
    Inherits CustomGesture

    Public Sub New()
        MyClass.New(Nothing)
    End Sub

    Public Sub New(ByVal strokeInfo As StrokeInfo)
        MyBase.New(strokeInfo)
        Name = "Circle"
    End Sub

    Protected Overrides Function Recognize() As Boolean
        Return _
            StrokeInfo.StrokeStatistics.StartEndProximity < CLOSED_PROXIMITY _
            AndAlso _
            ( _
                ( _
                    StrokeInfo.StrokeStatistics.StopPoints <> 4 _
                    AndAlso StrokeInfo.Stroke.PolylineCusps.Length < 4 _
                ) _
                OrElse StrokeInfo.StrokeStatistics.StopPoints >= 8 _
            ) _
            AndAlso StrokeInfo.StrokeStatistics.Square > 0.37 _
            AndAlso StrokeInfo.StrokeStatistics.Square < 0.63 _
            AndAlso StrokeInfo.IsMatch(Vectors.StartTick & Vectors.Rights & _
                Vectors.Downs & Vectors.Lefts & Vectors.Ups & Vectors.EndTick, 90, True, True)
    End Function

End Class
