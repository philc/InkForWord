Public Class Vectors
    Public Shared Start As String = "^"
    Public Shared [End] As String = "$"
    Public Shared Tick As String = "([A-Z]{1,2},){1,6}"
    Public Shared StartTick As String = Start & Tick
    Public Shared EndTick As String = Tick & [End]
    Public Shared Rights As String = "(R,|RU,|RD,)+"
    Public Shared Lefts As String = "(L,|LU,|LD,)+"
    Public Shared Ups As String = "(U,|LU,|RU,)+"
    Public Shared Downs As String = "(D,|LD,|RD,)+"
    Public Shared RightUps As String = "(RU,|U,|R,)+"
    Public Shared RightDowns As String = "(RD,|R,|D,)+"
    Public Shared LeftUps As String = "(LU,|L,|U,)+"
    Public Shared LeftDowns As String = "(LD,|L,|D,)+"
    Public Shared CornerTick As String = "(R,|RU,|RD,|D,|L,|LD,|U,|LU,){3}"
End Class

