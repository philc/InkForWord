Imports System.Drawing

Public Class Tick
    Inherits CustomGesture

    Public Sub New()
        MyClass.New(Nothing)
    End Sub

    Public Sub New(ByVal strokeInfo As StrokeInfo)
        MyBase.New(strokeInfo)
        Name = "Tick"
    End Sub

    Protected Overrides Function Recognize() As Boolean
        
        Return _
            StrokeInfo.IsMatch(Vectors.Downs) _
            AndAlso StrokeInfo.Stroke.PacketCount < 35
            
    End Function

End Class
