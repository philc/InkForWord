Imports Microsoft.Ink
Imports System.Drawing

Public Class Transpose
    Inherits CustomGesture

    Public Sub New()
        MyClass.New(Nothing)
    End Sub

    Public Sub New(ByVal strokeInfo As StrokeInfo)
        MyBase.New(strokeInfo)
        Name = "Transpose"
    End Sub

    Protected Overrides Function Recognize() As Boolean
        Dim startPoint As Point = StrokeInfo.Stroke.GetPoint(0)
        Dim endPoint As Point = StrokeInfo.Stroke.GetPoint(StrokeInfo.Stroke.PacketCount - 1)

        Console.WriteLine(startPoint.Y)
        Console.WriteLine(endPoint.Y)
        Console.WriteLine("Sub: " + (startPoint.Y - endPoint.Y).ToString())

        Return _
            StrokeInfo.IsMatch(Vectors.StartTick & Vectors.Ups & _
                Vectors.Rights & Vectors.Downs & Vectors.Rights & Vectors.Ups & _
                Vectors.EndTick) _
            AndAlso StrokeInfo.StrokeStatistics.Square <= 0.87 _
            AndAlso _
            ( _
                startPoint.Y - endPoint.Y < 900 _
                AndAlso startPoint.Y - endPoint.Y > -900 _
            )
    End Function

End Class
