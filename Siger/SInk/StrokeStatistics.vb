Imports System.ComponentModel

<TypeConverter(GetType(ExpandableObjectConverter))> _
Public Class StrokeStatistics
    <CLSCompliant(False)> Public _right As Double
    <Category("Direction Ratio")> _
    Public Property Right() As Double
        Get
            Return _right
        End Get
        Set(ByVal Value As Double)
            _right = Value
        End Set
    End Property

    <CLSCompliant(False)> Public _left As Double
    <Category("Direction Ratio")> _
    Public Property Left() As Double
        Get
            Return _left
        End Get
        Set(ByVal Value As Double)
            _left = Value
        End Set
    End Property

    <CLSCompliant(False)> Public _up As Double
    <Category("Direction Ratio")> _
    Public Property Up() As Double
        Get
            Return _up
        End Get
        Set(ByVal Value As Double)
            _up = Value
        End Set
    End Property

    <CLSCompliant(False)> Public _down As Double
    <Category("Direction Ratio")> _
    Public Property Down() As Double
        Get
            Return _down
        End Get
        Set(ByVal Value As Double)
            _down = Value
        End Set
    End Property

    <CLSCompliant(False)> Public _rightUp As Double
    <Category("Direction Ratio")> _
    Public Property RightUp() As Double
        Get
            Return _rightUp
        End Get
        Set(ByVal Value As Double)
            _rightUp = Value
        End Set
    End Property

    <CLSCompliant(False)> Public _rightDown As Double
    <Category("Direction Ratio")> _
    Public Property RightDown() As Double
        Get
            Return _rightDown
        End Get
        Set(ByVal Value As Double)
            _rightDown = Value
        End Set
    End Property

    <CLSCompliant(False)> Public _leftUp As Double
    <Category("Direction Ratio")> _
    Public Property LeftUp() As Double
        Get
            Return _leftUp
        End Get
        Set(ByVal Value As Double)
            _leftUp = Value
        End Set
    End Property

    <CLSCompliant(False)> Public _leftDown As Double
    <Category("Direction Ratio")> _
    Public Property LeftDown() As Double
        Get
            Return _leftDown
        End Get
        Set(ByVal Value As Double)
            _leftDown = Value
        End Set
    End Property


    <CLSCompliant(False)> Public _square As Double
    <Category("Overall Ratio")> _
    Public Property Square() As Double
        Get
            Return _square
        End Get
        Set(ByVal Value As Double)
            _square = Value
        End Set
    End Property

    <CLSCompliant(False)> Public _diagonal As Double
    <Category("Overall Ratio")> _
    Public Property Diagonal() As Double
        Get
            Return _diagonal
        End Get
        Set(ByVal Value As Double)
            _diagonal = Value
        End Set
    End Property

    <CLSCompliant(False)> Public _stopPoints As Integer
    <Category("Other Statistics")> _
    Public Property StopPoints() As Integer
        Get
            Return _stopPoints
        End Get
        Set(ByVal Value As Integer)
            _stopPoints = Value
        End Set
    End Property

    <CLSCompliant(False)> Public _startEndProximity As Double
    <Category("Other Statistics")> _
    Public Property StartEndProximity() As Double
        Get
            Return _startEndProximity
        End Get
        Set(ByVal Value As Double)
            _startEndProximity = Value
        End Set
    End Property

    Public Sub New()

    End Sub

End Class
