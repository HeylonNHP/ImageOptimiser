Option Explicit On

Imports System.Runtime.InteropServices

Public Class visibleScrollbars
    Private Const GWL_STYLE As Integer = -16
    Private Const WS_HSCROLL = &H100000
    Private Const WS_VSCROLL = &H200000

    <DllImport("user32.dll", SetLastError:=True)> _
    Private Shared Function GetWindowLong(ByVal hWnd As IntPtr,
                       ByVal nIndex As Integer) As Integer
    End Function

    ' sometimes you use wrappers since many, many, many things could call
    ' SendMessage and so that your code doesnt need to know all the MSG params
    Friend Shared Function IsVScrollVisible(ByVal ctl As Control) As Boolean
        Dim wndStyle As Integer = GetWindowLong(ctl.Handle, GWL_STYLE)
        Return ((wndStyle And WS_VSCROLL) <> 0)

    End Function
    Friend Shared Function IsHScrollVisible(ByVal ctl As Control) As Boolean
        Dim wndStyle As Integer = GetWindowLong(ctl.Handle, GWL_STYLE)
        Return ((wndStyle And WS_HSCROLL) <> 0)

    End Function
End Class
