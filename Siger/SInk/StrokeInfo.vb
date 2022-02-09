Imports Microsoft.Ink
Imports System.Text.RegularExpressions
Imports System.Text
Imports System.Drawing
Imports System.ComponentModel

Public Class StrokeInfo

    <CLSCompliant(False)> Protected _distances As Distances
    <CLSCompliant(False)> Protected _stroke As Stroke
    <CLSCompliant(False)> Protected _vectors As String
    <CLSCompliant(False)> Protected _strokeStatistics As New StrokeStatistics
    <CLSCompliant(False)> Protected _penVelocity() As Double


    Public Sub New(ByVal stroke As Stroke)
        Me.Stroke = stroke
        _vectors = GetVectors(stroke)
        _strokeStatistics = GetStrokeStatistics()
    End Sub

    <Browsable(True), _
    Category("General")> _
    Public ReadOnly Property Vectors() As String
        Get
            Return _vectors
        End Get
    End Property

    <Browsable(True), _
    Category("General")> _
    Public ReadOnly Property PenVelocity() As String
        Get
            Dim sb As New StringBuilder
            For Each d As Double In _penVelocity
                sb.Append(d).Append(vbCrLf)
            Next
            Return sb.ToString()
        End Get
    End Property

    <Browsable(True), _
    Category("Stroke Statistics")> _
    Public ReadOnly Property StrokeStatistics() As StrokeStatistics
        Get
            Return _strokeStatistics
        End Get
    End Property

    <Browsable(False)> _
    Public Property Stroke() As Stroke
        Get
            Return _stroke
        End Get
        Set(ByVal Value As Stroke)
            _stroke = Value
        End Set
    End Property

    Public Function IsMatch(ByVal matchPattern As String, ByVal rotationAngle As Double, ByVal flipX As Boolean, ByVal flipY As Boolean) As Boolean
        Dim pattern As New Regex(matchPattern)
        Dim s As Stroke = Stroke.Ink.Clone().Strokes(0)
        Dim v As String = _vectors

        Console.WriteLine(v)
        Console.WriteLine(matchPattern)

        If rotationAngle = 0 Then rotationAngle = 360
        Dim str As String
        str = ""
        For i As Integer = 1 To 360 / rotationAngle
            If pattern.IsMatch(v) Then
                Return True
            End If
            If rotationAngle = 90 Then
                If str = "" Then str = matchPattern
                str = str.Replace("U", "X").Replace("L", "U").Replace("D", "L").Replace("R", "D").Replace("X", "R").Replace("DR", "RD").Replace("DL", "LD").Replace("UR", "RU").Replace("UL", "LU")
                pattern = New Regex(str)
            Else
                s.Rotate(rotationAngle, New Point(0, 0))
                v = GetVectors(s)
            End If
        Next
        s.Ink.Dispose()

        If flipX Then
            s = Stroke.Ink.Clone().Strokes(0)
            Dim p() As Point = Stroke.GetPoints()
            Dim max As Integer = s.GetBoundingBox().Width
            For i As Integer = 0 To p.Length - 1
                p(i).X = max - p(i).x
            Next
            s.SetPoints(p)
            v = GetVectors(s)
            If pattern.IsMatch(v) Then
                Return True
            End If
        End If

        If flipY Then
            s = Stroke.Ink.Clone().Strokes(0)
            Dim p() As Point = Stroke.GetPoints()
            Dim max As Integer = s.GetBoundingBox().Height
            For i As Integer = 0 To p.Length - 1
                p(i).y = max - p(i).y
            Next
            s.SetPoints(p)
            v = GetVectors(s)
            If pattern.IsMatch(v) Then
                Return True
            End If
        End If

        Return False
    End Function

    Public Function IsMatch(ByVal pattern As String) As Boolean
        Return IsMatch(pattern, 0, False, False)
    End Function

    Protected Function GetStrokeStatistics() As StrokeStatistics
        Dim strokeStats As New StrokeStatistics
        With strokeStats
            .Down = _distances.Down / _distances.Total
            .Up = _distances.Up / _distances.Total
            .Right = _distances.Right / _distances.Total
            .Left = _distances.Left / _distances.Total
            .RightUp = _distances.RightUp / _distances.Total
            .RightDown = _distances.RightDown / _distances.Total
            .LeftUp = _distances.LeftUp / _distances.Total
            .LeftDown = _distances.LeftDown / _distances.Total
            .Square = .Left + .Right + .Up + .Down
            .Diagonal = .LeftDown + .LeftUp + .RightDown + .RightUp
        End With

        If (Stroke.PacketCount >= 11) Then
            strokeStats.StopPoints = CalculateStopPoints(Stroke)
        Else
            strokeStats.StopPoints = 0
        End If

        Dim startPoint As Point = Stroke.GetPoint(0)
        Dim endPoint As Point = Stroke.GetPoint(Stroke.PacketCount - 1)

        strokeStats.StartEndProximity = Math.Sqrt((startPoint.X - endPoint.X) ^ 2 + (startPoint.Y - endPoint.Y) ^ 2)

        Return strokeStats
    End Function

    Private Function CalculateStopPoints(ByVal stroke As Stroke) As Integer
        Dim stopPoints As Integer
        Dim lastPoint As Point
        ' Dim distance As Double
        Dim first As Boolean = True
        Dim points() As Point = stroke.GetPoints()
        Dim distances(points.Length - 2) As Double
        For i As Integer = 0 To points.Length - 1
            Dim p As Point = points(i)
            If Not first Then
                distances(i - 1) = Math.Sqrt((Math.Abs(p.X - lastPoint.X) ^ 2) + (Math.Abs(p.Y - lastPoint.Y) ^ 2))
            Else
                first = False
            End If
            lastPoint = p
        Next

        Dim movingAverage(distances.Length - 11) As Double
        Dim average As Double
        For i As Integer = 0 To 9
            average += distances(i)
        Next

        Dim max As Double
        For i As Integer = 0 To distances.Length - 11
            movingAverage(i) = average / 10
            average = average - distances(i) + distances(i + 10)
            If movingAverage(i) > max Then max = movingAverage(i)
        Next

        Dim isStopped As Boolean = False
        Dim stopPoint As Double
        For i As Integer = 0 To movingAverage.Length - 1
            If (i = 0 OrElse movingAverage(i) <= movingAverage(i - 1)) And _
                (i = movingAverage.Length - 1 OrElse movingAverage(i) < movingAverage(i + 1)) And _
                Not isStopped Then
                stopPoints += 1
                stopPoint = movingAverage(i)
                isStopped = True
            End If
            If movingAverage(i) > stopPoint + (max * 0.1) Then
                isStopped = False
            End If
            If i > (movingAverage.Length * 0.15) And stopPoints = 0 Then
                stopPoints = 1
            End If
        Next
        If Not isStopped Then stopPoints += 1

        _penVelocity = movingAverage

        Return stopPoints
    End Function

    Public Function MaxDifference(ByVal vps() As Double) As Double
        Dim maxDiff As Double

        For i As Integer = 0 To vps.Length - 2
            For j As Integer = 1 To vps.Length - 1
                If Math.Abs(vps(i) - vps(j)) > maxDiff Then maxDiff = Math.Abs(vps(i) - vps(j))
            Next
        Next

        Return maxDiff
    End Function

    Protected Function GetVectors(ByVal stroke As Stroke) As String
        _distances = New Distances

        Dim previous As Point = Nothing
        Dim current As Point = Nothing
        Dim isFirst As Boolean = True

        Dim vectors As New StringBuilder

        Const MIN_DISTANCE As Integer = 60

        For Each p As Point In stroke.GetPoints()
            If isFirst Then
                previous = p
                isFirst = False
            Else
                If Math.Sqrt(CDbl(p.X - previous.X) ^ 2 + CDbl(p.Y - previous.Y) ^ 2) >= MIN_DISTANCE Then
                    Dim xDif As Integer = p.X - previous.X
                    Dim yDif As Integer = p.Y - previous.Y

                    Dim xyRatio As Double = Math.Abs(CDbl(xDif) / CDbl(yDif))
                    If xyRatio >= 0.5 And xyRatio <= 2.0 Then
                        If xDif >= 0 And yDif >= 0 Then
                            vectors.Append("RD,")
                            _distances.RightDown += Distance(xDif, yDif)
                        ElseIf xDif >= 0 And yDif < 0 Then
                            vectors.Append("RU,")
                            _distances.RightUp += Distance(xDif, yDif)
                        ElseIf xDif < 0 And yDif >= 0 Then
                            vectors.Append("LD,")
                            _distances.LeftDown += Distance(xDif, yDif)
                        ElseIf xDif < 0 And yDif < 0 Then
                            vectors.Append("LU,")
                            _distances.RightDown += Distance(xDif, yDif)
                        End If
                    Else
                        If Math.Abs(xDif) >= Math.Abs(yDif) Then
                            If xDif > 0 Then
                                vectors.Append("R,")
                                _distances.Right += Distance(xDif, yDif)
                            Else
                                vectors.Append("L,")
                                _distances.Left += Distance(xDif, yDif)
                            End If
                        Else
                            If yDif > 0 Then
                                vectors.Append("D,")
                                _distances.Down += Distance(xDif, yDif)
                            Else
                                vectors.Append("U,")
                                _distances.Up += Distance(xDif, yDif)
                            End If
                        End If
                    End If
                    previous = p
                End If
            End If
        Next

        With _distances
            .Total = .Right + .Left + .Up + .Down + .RightUp + .RightDown + .LeftUp + .LeftDown
        End With

        Return vectors.ToString()

    End Function

    Protected Function Distance(ByVal x As Double, ByVal y As Double) As Double
        Return Math.Sqrt(x * x + y * y)
    End Function


End Class
