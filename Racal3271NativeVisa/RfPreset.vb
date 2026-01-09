Option Strict On
Option Explicit On

Public Class RfPreset
    Public Property Name As String
    Public Property FrequencyHz As Double
    Public Property LevelDbm As Double
    Public Property RfOn As Boolean = True

    Public Overrides Function ToString() As String
        Return Name
    End Function
End Class
