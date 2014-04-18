Imports System.Runtime.InteropServices

Public Class Win32API
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function GetForegroundWindow() As IntPtr
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
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)> _
    Public Shared Function GetKeyState(ByVal nVirtKey As Keys) As Short
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function GetKeyboardState(ByVal keyState() As Byte) As Boolean
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

End Class
