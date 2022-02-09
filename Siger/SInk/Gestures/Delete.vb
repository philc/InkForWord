Imports Microsoft.Ink
Imports System.Text.RegularExpressions

Public Class Delete
    Inherits CustomGesture

    Public Sub New()
        MyClass.New(Nothing)
    End Sub

    Public Sub New(ByVal strokeInfo As StrokeInfo)
        MyBase.New(strokeInfo)
        Name = "Delete"
    End Sub

    Protected Overrides Function Recognize() As Boolean
        Return _
            StrokeInfo.StrokeStatistics.StartEndProximity > 400 _
            AndAlso StrokeInfo.IsMatch(Vectors.StartTick + Vectors.RightUps + _
                Vectors.LeftUps + Vectors.LeftDowns + Vectors.RightDowns + _
                Vectors.RightUps + "(LU,|L,|U,)*" & Vectors.End, 0, False, False)
    End Function

End Class
