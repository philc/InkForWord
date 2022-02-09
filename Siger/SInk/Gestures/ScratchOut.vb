Imports Microsoft.Ink
Imports System.Text.RegularExpressions

Public Class ScratchOut
    Inherits CustomGesture

    Public Sub New()
        MyClass.New(Nothing)
    End Sub

    Public Sub New(ByVal strokeInfo As StrokeInfo)
        MyBase.New(strokeInfo)
        Name = "ScratchOut"
    End Sub


    Protected Overrides Function Recognize() As Boolean
        ' Look for at least 3 rights and lefts
        Return _
            StrokeInfo.StrokeStatistics.Right + StrokeInfo.StrokeStatistics.Left > 0.9 _
            AndAlso StrokeInfo.IsMatch("(" + Vectors.Rights + ".*" + Vectors.Lefts + "){2}")

    End Function


End Class
