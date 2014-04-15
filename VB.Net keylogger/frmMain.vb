Public Class frmMain
    Private WithEvents WindowsHook As New Hook

    Private Sub WindowsHook_KeyboardChange(nCode As Integer, wParam As Hook.WM_KEYBOARD_MSG, ByRef lParam As Hook.KBDLLHOOKSTRUCT, ByRef cancel As Boolean) Handles WindowsHook.KeyboardChange

    End Sub
End Class
