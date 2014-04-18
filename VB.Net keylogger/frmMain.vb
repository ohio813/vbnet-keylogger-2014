Public Class frmMain
    Private WithEvents WinHook As New Hook
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        WinHook.Hook(Hook.HookType.WH_KEYBOARD_LL)
    End Sub
    Private Sub frmMain_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        WinHook.UnHook(Hook.HookType.WH_KEYBOARD_LL)
    End Sub
    Private Shared builder As New System.Text.StringBuilder
    Private Sub WinHook_KeyboardChange(nCode As Integer, wParam As Hook.WM_KEYBOARD_MSG, ByRef lParam As Hook.KBDLLHOOKSTRUCT, ByRef cancel As Boolean) Handles WinHook.KeyboardChange
        Debug.Print("{5}>Hook_KeyboardChange: nCode={0}, wParam={1}, vkCode={2}, scanCode={3}, flags={4}, dwExtraInfo={6}", nCode, wParam.ToString, lParam.vkCode, lParam.scanCode, lParam.flags, DateTime.Now.AddTicks(lParam.time), lParam.dwExtraInfo)
        If nCode >= 0 Then
            If wParam = Hook.WM_KEYBOARD_MSG.WM_KEYDOWN Or wParam = Hook.WM_KEYBOARD_MSG.WM_SYSKEYDOWN Then
                Dim dwThreadID = Win32API.GetWindowThreadProcessId(Win32API.GetForegroundWindow, vbNull)
                Dim keyblayoutID As Integer = Win32API.GetKeyboardLayout(dwThreadID)
                Dim ScanCode As Integer = Win32API.MapVirtualKeyEx(lParam.vkCode, 2, keyblayoutID)

                Dim KeyStates(255) As Byte
                Dim result As Boolean = Win32API.GetKeyboardState(KeyStates)
                Dim FinalChar As New System.Text.StringBuilder(4)

                Dim ret = Win32API.ToUnicodeEx(lParam.vkCode, ScanCode, KeyStates, FinalChar, FinalChar.Capacity, lParam.flags, keyblayoutID)
                Dim vkCode As Integer
                If ret = 1 Then
                    vkCode = AscW(FinalChar.ToString)
                Else
                    vkCode = lParam.vkCode
                End If

                Select Case vkCode
                    'If we can block events
                    Case My.Computer.Keyboard.CtrlKeyDown And Keys.Escape
                        Me.Text = ("Ctrl + Esc blocked")
                        cancel = True
                    Case My.Computer.Keyboard.AltKeyDown And Keys.Escape
                        Me.Text = ("Alt + Escape blocked")
                        cancel = True
                    Case My.Computer.Keyboard.AltKeyDown And Keys.Tab
                        Me.Text = ("Alt + Tab blockd")
                        cancel = True  'Block event
                    Case My.Computer.Keyboard.CtrlKeyDown And My.Computer.Keyboard.AltKeyDown And Keys.Delete
                        Me.Text = ("Ctrl + Alt + Delete blocked")
                        ' This needs investigation is not working on win8 event no cancel
                        cancel = True  'Block event
                    Case My.Computer.Keyboard.CtrlKeyDown And My.Computer.Keyboard.AltKeyDown And Keys.S
                        Me.Visible = Not Me.Visible 'Hide/Show form
                        cancel = True  'Block event
                    Case Keys.CapsLock, Keys.RShiftKey, Keys.LShiftKey, Keys.RControlKey, Keys.LControlKey, Keys.LMenu, Keys.RMenu
                    Case Keys.Return
                        builder.Append("<" & Keys.Return.ToString & ">")
                    Case Keys.Back
                        builder.Append("<" & Keys.Back.ToString & ">")
                    Case Keys.Tab
                        builder.Append("<" & Keys.Tab.ToString & ">")
                    Case Keys.Escape
                        builder.Append("<" & Keys.Escape.ToString & ">")
                    Case Else
                        Dim keysDown As String = IIf(My.Computer.Keyboard.CtrlKeyDown, Keys.Control.ToString & " ", "")
                        keysDown &= IIf(My.Computer.Keyboard.AltKeyDown, Keys.Alt.ToString & " ", "")
                        keysDown &= IIf(My.Computer.Keyboard.ShiftKeyDown, Keys.Shift.ToString & " ", "")
                        If keysDown.Length > 0 And keysDown <> (Keys.Shift.ToString & " ") Then
                            builder.Append("<" & keysDown & lParam.vkCode.ToString & ">")
                        Else
                            If ret = 1 Then
                                builder.Append(ChrW(vkCode))
                            Else
                                builder.Append("<" & lParam.vkCode.ToString & ">")
                            End If
                        End If
                End Select
            End If
        End If
        RichTextBox1.Text = builder.ToString
    End Sub
End Class
