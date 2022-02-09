Imports Microsoft.Ink
Imports System.Text.RegularExpressions

Public Class HorizontalLine
    Inherits CustomGesture

    Public Sub New()
        MyClass.New(Nothing)
    End Sub

    Public Sub New(ByVal strokeInfo As StrokeInfo)
        MyBase.New(strokeInfo)
        Name = "HorizontalLine"
    End Sub

    Protected Overrides Function Recognize() As Boolean
        Return _
            StrokeInfo.IsMatch(Vectors.StartTick + Vectors.Rights + Vectors.EndTick) _
            AndAlso StrokeInfo.StrokeStatistics.Right > 0.95
    End Function

End Class