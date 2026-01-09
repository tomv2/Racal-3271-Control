Option Strict On
Option Explicit On

Imports System
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Collections.Generic

' Native VISA (visa32.dll) message-based I/O. Mirrors NI's example:
' viOpenDefaultRM -> viOpen -> viSetAttribute(timeout) -> viWrite/viRead -> viClose.

Public Module VisaNative
    Public Const VI_SUCCESS As Integer = 0
    Public Const VI_NULL As Integer = 0

    ' Timeout attribute (ms)
    Public Const VI_ATTR_TMO_VALUE As Integer = &H3FFF001A
    Public Const VI_FIND_BUFLEN As Integer = 256

    <DllImport("visa32.dll", CallingConvention:=CallingConvention.StdCall)>
    Public Function viOpenDefaultRM(ByRef sesn As IntPtr) As Integer
    End Function

    <DllImport("visa32.dll", CallingConvention:=CallingConvention.StdCall, CharSet:=CharSet.Ansi)>
    Public Function viOpen(ByVal rmSesn As IntPtr,
                           ByVal name As String,
                           ByVal mode As Integer,
                           ByVal timeout As Integer,
                           ByRef vi As IntPtr) As Integer
    End Function

    <DllImport("visa32.dll", CallingConvention:=CallingConvention.StdCall)>
    Public Function viSetAttribute(ByVal vi As IntPtr, ByVal attrName As Integer, ByVal attrValue As Integer) As Integer
    End Function

    <DllImport("visa32.dll", CallingConvention:=CallingConvention.StdCall)>
    Public Function viWrite(ByVal vi As IntPtr, ByVal buf() As Byte, ByVal cnt As Integer, ByRef retCnt As Integer) As Integer
    End Function

    <DllImport("visa32.dll", CallingConvention:=CallingConvention.StdCall)>
    Public Function viRead(ByVal vi As IntPtr, ByVal buf() As Byte, ByVal cnt As Integer, ByRef retCnt As Integer) As Integer
    End Function

    <DllImport("visa32.dll", CallingConvention:=CallingConvention.StdCall)>
    Public Function viClose(ByVal vi As IntPtr) As Integer
    End Function

    <DllImport("visa32.dll", CallingConvention:=CallingConvention.StdCall, CharSet:=CharSet.Ansi)>
    Public Function viFindRsrc(ByVal rmSesn As IntPtr,
                               ByVal expr As String,
                               ByRef findList As IntPtr,
                               ByRef retCnt As Integer,
                               ByVal desc As StringBuilder) As Integer
    End Function

    <DllImport("visa32.dll", CallingConvention:=CallingConvention.StdCall, CharSet:=CharSet.Ansi)>
    Public Function viFindNext(ByVal findList As IntPtr, ByVal desc As StringBuilder) As Integer
    End Function
End Module

Public Class VisaMessageSession
    Implements IDisposable

    Private rm As IntPtr = IntPtr.Zero
    Private instr As IntPtr = IntPtr.Zero
    Private disposedValue As Boolean

    Public ReadOnly Property IsOpen As Boolean
        Get
            Return instr <> IntPtr.Zero
        End Get
    End Property

    Public Sub Open(resource As String, Optional timeoutMs As Integer = 5000)
        Dim st As Integer = viOpenDefaultRM(rm)
        If st < VI_SUCCESS Then Throw New Exception("viOpenDefaultRM failed: " & st)

        st = viOpen(rm, resource, VI_NULL, VI_NULL, instr)
        If st < VI_SUCCESS Then
            viClose(rm) : rm = IntPtr.Zero
            Throw New Exception("viOpen failed: " & st & " for " & resource)
        End If

        ' Best-effort timeout
        viSetAttribute(instr, VI_ATTR_TMO_VALUE, timeoutMs)
    End Sub

    Public Sub WriteLine(cmd As String)
        If Not IsOpen Then Throw New InvalidOperationException("Session not open.")
        If Not cmd.EndsWith(Environment.NewLine, StringComparison.Ordinal) Then
            cmd &= Environment.NewLine
        End If

        Dim data As Byte() = Encoding.ASCII.GetBytes(cmd)
        Dim ret As Integer = 0
        Dim st As Integer = viWrite(instr, data, data.Length, ret)
        If st < VI_SUCCESS Then Throw New Exception("viWrite failed: " & st)
    End Sub

    Public Function ReadRaw(Optional maxBytes As Integer = 4096) As String
        If Not IsOpen Then Throw New InvalidOperationException("Session not open.")
        Dim buf(maxBytes - 1) As Byte
        Dim ret As Integer = 0
        Dim st As Integer = viRead(instr, buf, buf.Length, ret)
        If st < VI_SUCCESS Then Throw New Exception("viRead failed: " & st)
        Return Encoding.ASCII.GetString(buf, 0, Math.Max(0, ret))
    End Function

    Public Function QueryLine(cmd As String, Optional maxBytes As Integer = 4096) As String
        WriteLine(cmd)
        Return ReadRaw(maxBytes).Trim()
    End Function

    Public Shared Function FindInstruments(Optional expr As String = "?*INSTR") As List(Of String)
        Dim results As New List(Of String)()

        Dim rmSesn As IntPtr = IntPtr.Zero
        Dim st As Integer = viOpenDefaultRM(rmSesn)
        If st < VI_SUCCESS Then Return results

        Dim findList As IntPtr = IntPtr.Zero
        Dim retCnt As Integer = 0
        Dim sb As New StringBuilder(VI_FIND_BUFLEN)

        st = viFindRsrc(rmSesn, expr, findList, retCnt, sb)
        If st >= VI_SUCCESS AndAlso retCnt > 0 Then
            results.Add(sb.ToString())
            For i As Integer = 2 To retCnt
                sb.Clear()
                sb.EnsureCapacity(VI_FIND_BUFLEN)
                st = viFindNext(findList, sb)
                If st < VI_SUCCESS Then Exit For
                results.Add(sb.ToString())
            Next
            viClose(findList)
        End If

        viClose(rmSesn)
        Return results
    End Function

    Public Sub Close()
        If instr <> IntPtr.Zero Then
            viClose(instr)
            instr = IntPtr.Zero
        End If
        If rm <> IntPtr.Zero Then
            viClose(rm)
            rm = IntPtr.Zero
        End If
    End Sub

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            Close()
            disposedValue = True
        End If
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(disposing:=True)
        GC.SuppressFinalize(Me)
    End Sub
End Class
