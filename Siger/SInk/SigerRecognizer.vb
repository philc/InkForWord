Imports Microsoft.Ink

Public Class SigerRecognizer
    Private _recognizerList As ArrayList

    Public Sub New()
        _recognizerList = New ArrayList
    End Sub

    Public Property RecognizerList() As ArrayList
        Get
            Return _recognizerList
        End Get
        Set(ByVal Value As ArrayList)
            _recognizerList = Value
        End Set
    End Property

    Public Function Recognize(ByVal stroke As Stroke) As CustomGesture()
        Try
            Dim strokeInfo As New StrokeInfo(stroke)
            Dim resultsList As New ArrayList
            For Each gesture As CustomGesture In _recognizerList
                gesture.StrokeInfo = strokeInfo
                If gesture.IsMatch Then
                    resultsList.Add(gesture)
                End If
            Next

            Dim maxResultIndex As Integer = resultsList.Count - 1
            Dim resultArray(maxResultIndex) As CustomGesture
            For i As Integer = 0 To maxResultIndex
                resultArray(i) = resultsList(i)
            Next

            Return resultArray

        Catch ex As Exception
            Throw ex
        End Try
    End Function

End Class
