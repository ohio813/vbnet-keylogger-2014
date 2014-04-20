Imports System.Runtime.InteropServices

Public Class Win32API
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function GetForegroundWindow() As IntPtr
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto, ExactSpelling:=True)>
    Public Shared Function WindowFromPoint(ByVal Point As Drawing.Point) As IntPtr
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto, ExactSpelling:=True)>
    Public Shared Function WindowFromPoint(ByVal xPoint As Integer, ByVal yPoint As Integer) As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function GetWindowThreadProcessId(ByVal hwnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
    End Function
#Region "Keyboard API"
    <DllImport("user32.dll", CharSet:=CharSet.Auto, ExactSpelling:=True)>
    Public Shared Function GetKeyboardLayout(ByVal dwLayout As Integer) As Integer
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function MapVirtualKeyEx(ByVal uCode As Integer, ByVal nMapType As Integer, ByVal dwhkl As Integer) As Integer
    End Function
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Shared Function GetKeyState(ByVal nVirtKey As Keys) As Short
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function GetKeyboardState(ByVal keyState() As Byte) As Long
    End Function
#End Region

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function ToUnicodeEx(wVirtKey As UInteger,
                                wScanCode As UInteger,
                                lpKeyState As Byte(),
                                <Out()>
                                <MarshalAs(UnmanagedType.LPWStr, SizeConst:=64)>
                                ByVal lpChar As System.Text.StringBuilder,
                                cchBuff As Integer,
                                wFlags As UInteger,
                                dwhkl As IntPtr) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function GetWindowTextLength(ByVal hwnd As IntPtr) As Integer
    End Function
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Ansi)> _
    Private Shared Function GetWindowText(hwnd As IntPtr, <MarshalAs(UnmanagedType.LPStr)> lpString As System.Text.StringBuilder, cch As Integer) As Integer
    End Function
    Public Shared Function GetWindowText(hwnd As Integer) As String
        Dim length As Integer = GetWindowTextLength(hwnd)
        Dim sb As New System.Text.StringBuilder("", length)

        GetWindowText(hwnd, sb, sb.Capacity + 1)

        Return sb.ToString
    End Function
End Class
