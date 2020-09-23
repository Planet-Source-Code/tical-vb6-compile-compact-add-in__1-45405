VERSION 5.00
Begin VB.Form frmAddIn 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Compacting ..."
   ClientHeight    =   840
   ClientLeft      =   2175
   ClientTop       =   1890
   ClientWidth     =   4290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Your application is being compacted. Please Wait!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public VBInstance As VBIDE.VBE      'VB IDE instance
Public Connect As Connect           'Connect Designer instance
Private WinDir As String

'This is executed when form loads
Private Sub Form_Load()

    On Error GoTo error_handle
        
    'Change mouse pointer to hourglass
    Screen.MousePointer = vbHourglass
    
    'Compile the application
    VBInstance.ActiveVBProject.MakeCompiledFile
    
    'Show my form
    Me.Show
    
    'Force OS to do events before compacting
    DoEvents
    
    'Compact it with UPX
    Call ExecuteAndWait("upx.exe --best --crp-ms=999999 --force --no-backup " & _
                         QT & VBInstance.ActiveVBProject.BuildFileName & QT)
    
    'Change label and form caption
    lblStatus.Caption = "Your application is being scrambled. Please Wait!"
    Me.Caption = "Scrambling ..."
    
    'Force OS to show changes
    DoEvents
    
    'Scramble the program
    Call UPX_Scramble(VBInstance.ActiveVBProject.BuildFileName)
    
    'Wait for 1 second so people can see the change in the form
    Sleep 1000
    
    'Change mouse pointer back to default arrow
    Screen.MousePointer = vbArrow
    
    'Hide the form
    Connect.Hide
    
    'Unload the form
    Unload Me
    
    Exit Sub

error_handle:
    
    If Err.Number = 91 Then
        MsgBox "Error. There must be at least one project open before you can use Compile & Compact Add-In.", vbExclamation, "VB 6 Compile & Compact error"
    ElseIf Err.Number < 0 Then
        MsgBox "Error. There is a syntax error in your code. Make sure your code compiles before you use Compile & Compact Add-In.", vbExclamation, "VB 6 Compile & Compact error"
    Else
        MsgBox "You did something stupid. Don't do it again!", vbExclamation, "VB 6 Compile & Compact error"
    End If

    'Hide the form
    Connect.Hide
    
    'Unload the form
    Unload Me
  
    Exit Sub

End Sub

'Execute a console program in background and wait for it to finish
Public Sub ExecuteAndWait(cmdline$)
    
    Dim NameOfProc As PROCESS_INFORMATION
    Dim NameStart As STARTUPINFO
    Dim X As Long
    
    'Get and set the length of STARTUPINFO structure
    NameStart.cb = Len(NameStart)
    
    'Create a detached process
    X = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, DETACHED_PROCESS, 0&, 0&, _
                       NameStart, NameOfProc)
    
    'Wait for it to finish
    X = WaitForSingleObject(NameOfProc.hProcess, INFINITE)
    
    'Close the process handle
    X = CloseHandle(NameOfProc.hProcess)
    
End Sub

Public Sub UPX_Scramble(filename$)
    Dim First As Byte, Second As Byte, Third As Byte, Fourth As Byte
    Dim sByte As Byte, i As Long
    Dim str As String
    
    'Initialise variables
    i = 0
    First = 0
    Second = 0
    Third = 0
    Fourth = 0
    
    'Open UPX'd application
    Open filename For Binary As 1
    
    'Loop byte by byte until the end of file
    Do While Not EOF(1)
    
        'Get byte
        Get 1, , sByte
    
        'Re-assign bytes
        First = Second
        Second = Third
        Third = Fourth
        Fourth = sByte
        
        'Combine a string we can check for bellow
        str = Chr$(First) & Chr$(Second) & Chr$(Third) & Chr$(Fourth)
        
        If str = "UPX0" Then
        
            'We found the first UPX symbol
            'so we replace it with "code"
            'in hex "code" = 63 6F 64 65
            Put 1, (i - 2), &H63
            Put 1, (i - 1), &H6F
            Put 1, (i + 0), &H64
            Put 1, (i + 1), &H65
                    
        ElseIf str = "UPX1" Then
        
            'We found the second UPX symbol
            'so we replace it with "text"
            'in hex "text" = 74 65 78 74
            Put 1, (i - 1), &H74
            Put 1, (i + 0), &H65
            Put 1, (i + 1), &H78
            Put 1, (i + 2), &H74
                    
        ElseIf str = "UPX!" Then
        
            'We found the last UPX symbol
            'replace this one nulls
            Put 1, (i - 5), &H0
            Put 1, (i - 4), &H0
            Put 1, (i - 3), &H0
            Put 1, (i - 2), &H0
            Put 1, (i - 1), &H0
            Put 1, (i + 0), &H0
            Put 1, (i + 1), &H0
            Put 1, (i + 2), &H0
            
            'Close the file
            Close 1
            
            'Exit the function
            Exit Sub
                    
        End If
        
        'increment byte counter
        i = i + 1
        
    Loop
 
End Sub

'Just in case the form loses focus, set it back
Private Sub Form_LostFocus()

    'Set the focus back on form
    Me.SetFocus
    
End Sub
