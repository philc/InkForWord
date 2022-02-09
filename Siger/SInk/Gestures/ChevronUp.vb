Imports System.Drawing

Public Class ChevronUp
    Inherits CustomGesture

    Public Sub New()
        MyClass.New(Nothing)
    End Sub

    Public Sub New(ByVal strokeInfo As StrokeInfo)
        MyBase.New(strokeInfo)
        Name = "ChevronUp"
    End Sub

    Protected Overrides Function Recognize() As Boolean
        Dim midPoint As Point = StrokeInfo.Stroke.GetPoint(StrokeInfo.Stroke.PacketCount / 2)
        Dim startPoint As Point = StrokeInfo.Stroke.GetPoint(0)
        Dim endPoint As Point = StrokeInfo.Stroke.GetPoint(StrokeInfo.Stroke.PacketCount - 1)

        Return _
            StrokeInfo.IsMatch(Vectors.RightUps & Vectors.RightDowns) _
            AndAlso midPoint.Y < startPoint.Y _
            AndAlso midPoint.Y < endPoint.Y _
            AndAlso StrokeInfo.StrokeStatistics.Right < 0.12 _
            AndAlso StrokeInfo.Vectors.StartsWith("RD") = False
    End Function

End Class
