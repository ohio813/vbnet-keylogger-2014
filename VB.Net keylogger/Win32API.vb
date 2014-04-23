Imports System.Runtime.InteropServices

Public Class Win32API

#Region "Keyboard Input Functions"
    ''' <summary>
    ''' Retrieves the active input locale identifier (formerly called the keyboard layout).
    ''' </summary>
    ''' <param name="idThread"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DllImport("user32.dll", CharSet:=CharSet.Auto, ExactSpelling:=True)> Public Shared Function GetKeyboardLayout(ByVal dwLayout As Integer) As IntPtr
    End Function
    ''' <summary>
    ''' Copies the status of the 256 virtual keys to the specified buffer. 
    ''' </summary>
    ''' <param name="keyState"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto, ExactSpelling:=True)> Public Shared Function GetKeyboardState(ByVal keyState() As Byte) As Boolean
    End Function
    ''' <summary>
    ''' Retrieves the status of the specified virtual key. 
    ''' The status specifies whether the key is up, down, or toggled (on, off—alternating each time the key is pressed). 
    ''' </summary>
    ''' <param name="nVirtKey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto, ExactSpelling:=True)> Public Shared Function GetKeyState(ByVal nVirtKey As Keys) As Short
    End Function
    ''' <summary>
    ''' Translates (maps) a virtual-key code into a scan code or character value, or translates a scan code into a virtual-key code. 
    ''' The function translates the codes using the input language and an input locale identifier.
    ''' </summary>
    ''' <param name="uCode"></param>
    ''' <param name="nMapType"></param>
    ''' <param name="dwhkl"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DllImport("user32.dll", CharSet:=CharSet.Auto)> Public Shared Function MapVirtualKeyEx(ByVal uCode As Integer, ByVal nMapType As Integer, ByVal dwhkl As Integer) As Integer
    End Function
    ''' <summary>
    ''' Translates the specified virtual-key code and keyboard state to the corresponding Unicode character or characters.
    ''' </summary>
    ''' <param name="wVirtKey"></param>
    ''' <param name="wScanCode"></param>
    ''' <param name="lpKeyState"></param>
    ''' <param name="lpChar"></param>
    ''' <param name="cchBuff"></param>
    ''' <param name="wFlags"></param>
    ''' <param name="dwhkl"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DllImport("user32.dll", CharSet:=CharSet.Auto)> Public Shared Function ToUnicodeEx(wVirtKey As UInteger,
                                                                                        wScanCode As UInteger,
                                                                                        lpKeyState As Byte(),
                                                                                        <Out()>
                                                                                        <MarshalAs(UnmanagedType.LPWStr, SizeConst:=64)>
                                                                                        ByVal lpChar As System.Text.StringBuilder,
                                                                                        cchBuff As Integer,
                                                                                        wFlags As UInteger,
                                                                                        dwhkl As IntPtr) As Integer
    End Function
#End Region

#Region "Cursor Functions"
    ''' <summary>
    ''' Moves the cursor to the specified screen coordinates. If the new coordinates are not within the screen rectangle set by the most recent ClipCursor function call,
    ''' the system automatically adjusts the coordinates so that the cursor stays within the rectangle. 
    ''' </summary>
    ''' <param name="X"></param>
    ''' <param name="Y"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DllImport("user32.dll", SetLastError:=True)> Public Shared Function SetCursorPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
    End Function
#End Region

#Region "Mouse Input Functions"
    ''' <summary>
    ''' Retrieves the current double-click time for the mouse. 
    ''' A double-click is a series of two clicks of the mouse button, the second occurring within a specified time after the first. 
    ''' The double-click time is the maximum number of milliseconds that may occur between the first and second click of a double-click. 
    ''' The maximum double-click time is 5000 milliseconds.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DllImport("user32.dll", CharSet:=CharSet.Auto, ExactSpelling:=True)> Public Shared Function GetDoubleClickTime() As Integer
    End Function
#End Region

#Region "Window Functions"
    ''' <summary>
    ''' Retrieves a handle to the foreground window (the window with which the user is currently working).
    ''' The system assigns a slightly higher priority to the thread that creates the foreground window than it does to other threads
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto, ExactSpelling:=True)> Public Shared Function GetForegroundWindow() As IntPtr
    End Function
    ''' <summary>
    ''' Retrieves the identifier of the thread that created the specified window and, optionally, the identifier of the process that created the window. 
    ''' </summary>
    ''' <param name="hwnd"></param>
    ''' <param name="lpdwProcessId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DllImport("user32.dll", SetLastError:=True)> Public Shared Function GetWindowThreadProcessId(ByVal hwnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
    End Function
    ''' <summary>
    ''' Copies the text of the specified window's title bar (if it has one) into a buffer. 
    ''' If the specified window is a control,the text of the control is copied. 
    ''' However,GetWindowText cannot retrieve the text of a control in another application.
    ''' </summary>
    ''' <param name="hwnd"></param>
    ''' <param name="lpString"></param>
    ''' <param name="cch"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DllImport("user32", SetLastError:=True, CharSet:=CharSet.Auto)> Public Shared Function GetWindowText(ByVal hWnd As IntPtr, <Out, MarshalAs(UnmanagedType.LPTStr)> ByVal lpString As System.Text.StringBuilder, ByVal nMaxCount As Integer) As Integer
    End Function
    ''' <summary>
    ''' Retrieves the length, in characters, of the specified window's title bar text (if the window has a title bar). 
    ''' If the specified window is a control, the function retrieves the length of the text within the control. 
    ''' However, GetWindowTextLength cannot retrieve the length of the text of an edit control in another application.
    ''' </summary>
    ''' <param name="hwnd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> Private Shared Function GetWindowTextLength(ByVal hwnd As IntPtr) As Integer
    End Function
    ''' <summary>
    ''' Retrieves a handle to the window that contains the specified point. 
    ''' </summary>
    ''' <param name="Point"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DllImport("user32.dll", CharSet:=CharSet.Auto, ExactSpelling:=True)> Public Shared Function WindowFromPoint(ByVal Point As Drawing.Point) As IntPtr
    End Function
    ''' <summary>
    ''' Retrieves a handle to the window that contains the specified point. 
    ''' </summary>
    ''' <param name="x"></param>
    ''' <param name="y"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DllImport("user32.dll", CharSet:=CharSet.Auto, ExactSpelling:=True)> Public Shared Function WindowFromPoint(ByVal x As Integer, ByVal y As Integer) As IntPtr
    End Function
    Public Shared Function GetWindowText(hwnd As Integer) As String
        Dim length As Integer = GetWindowTextLength(hwnd)
        Dim sb As New System.Text.StringBuilder(vbNullChar, length + 1)

        GetWindowText(hwnd, sb, sb.Capacity)

        Return sb.ToString
    End Function
#End Region
End Class
