Public Class frmMain
    Private WithEvents WinHook As New Hook
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Initialize Controls
        My.Settings.SessionKeyboardClick = 0
        My.Settings.SessionMouseMoves = 0
        My.Settings.SessionMouseClick = 0
        My.Settings.Save()
        WinHook.Hook(Hook.HookType.WH_KEYBOARD_LL)
        WinHook.Hook(Hook.HookType.WH_MOUSE_LL)
        Timer1.Enabled = True
    End Sub
    Private Sub frmMain_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        WinHook.UnHook(Hook.HookType.WH_KEYBOARD_LL)
        WinHook.UnHook(Hook.HookType.WH_MOUSE_LL)
    End Sub
    Private LastCheckedForegroundTitle As String = ""
    Private onStartKeyboardSessionClicks As Integer = My.Settings.TotalKeyboardClick
    Private Sub WinHook_KeyboardChange(nCode As Integer, wParam As Hook.WM_KEYBOARD_MSG, ByRef lParam As Hook.KBDLLHOOKSTRUCT, ByRef cancel As Boolean) Handles WinHook.KeyboardChange
        Debug.Print("{5}>KeyboardChange: nCode={0}, wParam={1}, vkCode={2}, scanCode={3}, flags={4}, dwExtraInfo={6}", nCode, wParam.ToString, lParam.vkCode, lParam.scanCode, lParam.flags, DateTime.Now.AddTicks(lParam.time), lParam.dwExtraInfo)
        My.Settings.SessionKeyboardClick = My.Settings.TotalKeyboardClick - onStartKeyboardSessionClicks
        My.Settings.Save()
        'Get current foreground window title
        Dim CurrentTitle = Win32API.GetWindowText(Win32API.GetForegroundWindow)
        'If title of the foreground window changed
        If CurrentTitle <> LastCheckedForegroundTitle Then
            LastCheckedForegroundTitle = CurrentTitle
            'Add a little header containing the application title
            Label1.Text = CurrentTitle
        End If

        If nCode >= 0 Then
            If wParam = Hook.WM_KEYBOARD_MSG.WM_KEYDOWN Or wParam = Hook.WM_KEYBOARD_MSG.WM_SYSKEYDOWN Then
                Dim dwThreadID = Win32API.GetWindowThreadProcessId(Win32API.GetForegroundWindow, vbNull)
                Dim keyblayoutID As Integer = Win32API.GetKeyboardLayout(dwThreadID)
                Dim ScanCode As Integer = Win32API.MapVirtualKeyEx(lParam.vkCode, 2, keyblayoutID)

                Dim CapitalState = CBool(Win32API.GetKeyState(Keys.Capital) And &H8000)
                Dim KeyStates(255) As Byte
                Win32API.GetKeyboardState(KeyStates)

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
                        RichTextBox1.AppendText("<" & Keys.Return.ToString & ">")
                    Case Keys.Back
                        RichTextBox1.AppendText("<" & Keys.Back.ToString & ">")
                    Case Keys.Tab
                        RichTextBox1.AppendText("<" & Keys.Tab.ToString & ">")
                    Case Keys.Escape
                        RichTextBox1.AppendText("<" & Keys.Escape.ToString & ">")
                    Case Else
                        Dim keysDown As String = IIf(My.Computer.Keyboard.CtrlKeyDown, Keys.Control.ToString & " ", "")
                        keysDown &= IIf(My.Computer.Keyboard.AltKeyDown, Keys.Alt.ToString & " ", "")
                        keysDown &= IIf(My.Computer.Keyboard.ShiftKeyDown, Keys.Shift.ToString & " ", "")
                        If keysDown.Length > 0 And keysDown <> (Keys.Shift.ToString & " ") Then
                            RichTextBox1.AppendText("<" & keysDown & lParam.vkCode.ToString & ">")
                        Else
                            If ret = 1 Then
                                RichTextBox1.AppendText(ChrW(vkCode))
                            Else
                                RichTextBox1.AppendText("<" & lParam.vkCode.ToString & ">")
                            End If
                        End If
                End Select
            End If
        End If
    End Sub
    Private onStartMouseSessionMoves As Integer = My.Settings.TotalMouseMoves
    Private onStartMouseSessionClicks As Integer = My.Settings.TotalMouseClick
    Private Sub WinHook_MouseChange(nCode As Integer, wParam As Hook.WM_MOUSE_MSG, lParam As Hook.MSLLHOOKSTRUCT, ByRef cancel As Boolean) Handles WinHook.MouseChange
        Debug.Print("{6}>MouseChange: nCode={0}, wParam={1}, x={2}, y={3}, mouseData={4}, flags={5}, dwExtraInfo={7}", nCode, wParam.ToString, lParam.pt.X, lParam.pt.Y, lParam.mouseData, lParam.flags, DateTime.Now.AddTicks(lParam.time), lParam.dwExtraInfo)
        My.Settings.SessionMouseMoves = My.Settings.TotalMouseMoves - onStartMouseSessionMoves
        My.Settings.SessionMouseClick = My.Settings.TotalMouseClick - onStartMouseSessionClicks
        My.Settings.Save()
        Label2.Text = Win32API.GetWindowText(Win32API.WindowFromPoint(lParam.pt))
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Stop()
        DataGridView1.Rows.Clear()
        For Each SettingsProperty As System.Configuration.SettingsProperty In My.Settings.Properties
            DataGridView1.Rows.Add(SettingsProperty.Name, My.Settings(SettingsProperty.Name))
        Next
        Timer1.Start()
    End Sub
End Class
