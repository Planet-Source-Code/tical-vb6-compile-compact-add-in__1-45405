Attribute VB_Name = "Module1"
'Module containing useful function definitions and constants

'CreateProcessA is used to execute a program
Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, _
               ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, _
               ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, _
               ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal _
               lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, _
               lpProcessInformation As PROCESS_INFORMATION) As Long
               
'WaitForSingleObject is used to wait for the process
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal _
               dwMilliseconds As Long) As Long
               
'Get windows directory
Declare Function GetWindowsDirectory Lib "kernel32" (ByVal lpBuffer As String, _
               ByVal nSize As Long) As Long
               
'Close handle terminates the connection the process we executed
Declare Function CloseHandle Lib "kernel32" (hObject As Long) As Boolean

'Sleep pauses the application for the number of milliseconds specified
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


'Constants used in form

'CreateProcessA function executes a console program in background
Global Const DETACHED_PROCESS = &H8&

'Wait infinitely for process to finish
Global Const INFINITE = -1&

'Quote string
Global Const QT As String = """"

'STARTUPINFO structure needed by the CreateProcessA function
Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

'PROCESS_INFORMATION structure needed by the CreateProcessA function
Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type
