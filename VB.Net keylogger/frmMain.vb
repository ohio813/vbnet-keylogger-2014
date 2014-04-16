Public Class frmMain
    Private WithEvents WinHook As New Hook
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        WinHook.Hook(Hook.HookType.WH_KEYBOARD_LL)
    End Sub
    Private Sub frmMain_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        WinHook.UnHook(Hook.HookType.WH_KEYBOARD_LL)
    End Sub
    Private Sub WinHook_KeyboardChange(nCode As Integer, wParam As Hook.WM_KEYBOARD_MSG, ByRef lParam As Hook.KBDLLHOOKSTRUCT, ByRef cancel As Boolean) Handles WinHook.KeyboardChange
        Debug.Print("{5}>Hook_KeyboardChange: nCode={0}, wParam={1}, vkCode={2}, scanCode={3}, flags={4}, dwExtraInfo={6}", nCode, wParam.ToString, lParam.vkCode, lParam.scanCode, lParam.flags, DateTime.Now.AddTicks(lParam.time), lParam.dwExtraInfo)

        If wParam = Hook.WM_KEYBOARD_MSG.WM_KEYDOWN Or wParam = Hook.WM_KEYBOARD_MSG.WM_SYSKEYDOWN Then
            'If we can block events
            If nCode >= 0 Then
                Select Case lParam.vkCode
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
                End Select
            End If
        End If

    End Sub
End Class
