Imports System.Runtime.InteropServices

Module Module1
    'Variable Declarations
    Private Const APPCOMMAND_VOLUME_MUTE As Integer = &H80000
    Private Const APPCOMMAND_VOLUME_UP As Integer = &HA0000
    Private Const APPCOMMAND_VOLUME_DOWN As Integer = &H90000
    Private Const WM_APPCOMMAND As Integer = &H319

    <DllImport("user32.dll")> Public Function SendMessageW(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function

    <DllImport("kernel32.dll")> Private Function GetConsoleWindow() As IntPtr
    End Function

    Sub Main()
        Dim intPtr As IntPtr = GetConsoleWindow()

        Console.WriteLine("Control Computer Volume With Following Options")
        Console.WriteLine("1. Volume Up")
        Console.WriteLine("2. Volume Down")
        Console.WriteLine("3. Mute / Unmute")
        Console.WriteLine("4. Exit")

        While (1)
            Dim input As Integer = Console.ReadLine()

            Select Case input
                Case 1
                    SendMessageW(intPtr, WM_APPCOMMAND, intPtr, New IntPtr(APPCOMMAND_VOLUME_UP))
                Case 2
                    SendMessageW(intPtr, WM_APPCOMMAND, intPtr, New IntPtr(APPCOMMAND_VOLUME_DOWN))
                Case 3
                    SendMessageW(intPtr, WM_APPCOMMAND, intPtr, New IntPtr(APPCOMMAND_VOLUME_MUTE))
                Case 4
                    Exit Sub
                Case Else
                    Console.WriteLine("Invalid Option")
            End Select
        End While

    End Sub

End Module
