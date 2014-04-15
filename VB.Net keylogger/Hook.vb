Imports System.Runtime.InteropServices
Public Class Hook
    Private Delegate Function CallBack(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
    Enum HookType As Integer
        WH_JOURNALRECORD = 0
        WH_JOURNALPLAYBACK = 1
        WH_KEYBOARD = 2
        WH_GETMESSAGE = 3
        WH_CALLWNDPROC = 4
        WH_CBT = 5
        WH_SYSMSGFILTER = 6
        WH_MOUSE = 7
        WH_HARDWARE = 8
        WH_DEBUG = 9
        WH_SHELL = 10
        WH_FOREGROUNDIDLE = 11
        WH_CALLWNDPROCRET = 12
        WH_KEYBOARD_LL = 13
        WH_MOUSE_LL = 14
    End Enum

    'Import for the SetWindowsHookEx function.
    <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)> _
    Private Overloads Shared Function SetWindowsHookEx _
          (ByVal HookType As HookType, ByVal HookProc As CallBack, _
           ByVal hInstance As IntPtr, ByVal wParam As Integer) As Integer
    End Function

    'Import for the CallNextHookEx function.
    <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)> _
    Private Overloads Shared Function CallNextHookEx _
          (ByVal idHook As Integer, ByVal nCode As Integer, _
           ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
    End Function
    'Import for the UnhookWindowsHookEx function.
    <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)> _
    Private Overloads Shared Function UnhookWindowsHookEx _
              (ByVal idHook As Integer) As Boolean
    End Function

    Private hHooks As New Dictionary(Of HookType, Integer)

#Region "WH_KEYBOARD_LL"
    Public Enum WM_KEYBOARD_MSG
        WM_KEYDOWN = &H100
        WM_KEYUP = &H101
        WM_SYSKEYDOWN = &H104
        WM_SYSKEYUP = &H105
    End Enum
    <Flags()> Public Enum KBDLLHOOKSTRUCTFlags As UInt32
        LLKHF_EXTENDED = &H1
        LLKHF_INJECTED = &H10
        LLKHF_ALTDOWN = &H20
        LLKHF_UP = &H80
    End Enum

    ' KeyboardStruct structure declaration.
    <StructLayout(LayoutKind.Sequential)> Public Structure KBDLLHOOKSTRUCT
        Public vkCode As VirtualKeysEnum.VirtualKeys
        Public scanCode As UInt32
        Public flags As KBDLLHOOKSTRUCTFlags
        Public time As UInt32
        Public dwExtraInfo As UIntPtr
    End Structure

    Public Event KeyboardChange(nCode As Integer, wParam As WM_KEYBOARD_MSG, ByRef lParam As KBDLLHOOKSTRUCT, ByRef cancel As Boolean)
    Private Function LowLevelKeyboardProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer

        Dim cancel As Boolean = False

        Dim myKeyboardHookStruct As New KBDLLHOOKSTRUCT()
        myKeyboardHookStruct = CType(Marshal.PtrToStructure(lParam, myKeyboardHookStruct.GetType()), KBDLLHOOKSTRUCT)

        RaiseEvent KeyboardChange(nCode, wParam, myKeyboardHookStruct, cancel)

        If cancel = True Then
            Return 1
        Else
            Return CallNextHookEx(hHooks(HookType.WH_KEYBOARD_LL), nCode, wParam, lParam)
        End If
    End Function
#End Region
End Class
