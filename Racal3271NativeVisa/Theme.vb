Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Windows.Forms

Public Class Theme
    Public Property Back As Color
    Public Property Panel As Color
    Public Property Text As Color
    Public Property Accent As Color
    Public Property InputBack As Color
End Class

Public Module Themes
    Public ReadOnly Light As New Theme With {
        .Back = Color.White,
        .Panel = Color.FromArgb(245, 246, 248),
        .Text = Color.FromArgb(30, 30, 30),
        .Accent = Color.FromArgb(0, 120, 215),
        .InputBack = Color.White
    }

    Public ReadOnly Dark As New Theme With {
        .Back = Color.FromArgb(30, 30, 30),
        .Panel = Color.FromArgb(45, 45, 48),
        .Text = Color.Gainsboro,
        .Accent = Color.FromArgb(0, 122, 204),
        .InputBack = Color.FromArgb(55, 55, 55)
    }
End Module
