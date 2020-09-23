VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   11745
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   14715
   _ExtentX        =   25956
   _ExtentY        =   20717
   _Version        =   393216
   Description     =   $"Connect.dsx":0000
   DisplayName     =   "VB 6 Compile & Compact"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public FormDisplayed As Boolean     'Visible or not
Public VBInstance As VBIDE.VBE      'VB IDE instance
Dim mfrmAddIn As New frmAddIn       'Our form instance

'VB IDE Menu & Command bar instance
Dim mcbMenuCommandBar As Office.CommandBarControl

'Command bar event handler
Public WithEvents MenuHandler As CommandBarEvents
Attribute MenuHandler.VB_VarHelpID = -1


'Hides our form
Sub Hide()
    
    'ignore any errors
    On Error Resume Next
    
    'not visible
    FormDisplayed = False
    
    'hide form itself
    mfrmAddIn.Hide
   
End Sub

'Shows our form
Sub Show()
  
    'ignore any errors
    On Error Resume Next
    
    'make sure our form instance is not null
    If mfrmAddIn Is Nothing Then
    
        'if it is then create a new instance
        Set mfrmAddIn = New frmAddIn
        
    End If
    
    'Connect the VB IDE instance of this Connect object to our forms'
    Set mfrmAddIn.VBInstance = VBInstance
    Set mfrmAddIn.Connect = Me
    'visible
    FormDisplayed = True
    'show the form
    mfrmAddIn.Show
   
End Sub


'This method adds the Add-In to VB
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    
    'contain any errors
    On Error GoTo error_handler
    
    'save the vb instance
    Set VBInstance = Application
    
    'If used by the wizard toolbar to start this wizard
    If ConnectMode = ext_cm_External Then
        
        'Show form
        Me.Show
        
    Else 'started at startup
        
        'sink the event
        Set mcbMenuCommandBar = AddToAddInCommandBar("Compile & Compact")
        Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
        
    End If
  
    
    'if it was started after startup
    If ConnectMode = ext_cm_AfterStartup Then
    
        'check if add-in wants to be displayed on connection
        If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
        
            'set this to display the form on connect
            Me.Show
            
        End If
        
    End If
  
    Exit Sub
    
'handle any errors
error_handler:
    
    MsgBox Err.Description
    
End Sub

'This method removes the Add-In from VB
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    
    'ignore any errors
    On Error Resume Next
    
    'delete the command bar entry
    mcbMenuCommandBar.Delete
   
    'shut down the Add-In
    If FormDisplayed Then
    
        'save the loading settings
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
        'not visible
        FormDisplayed = False
        
    Else 'not visible
    
        'just save loading settings
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
        
    End If
    
    'unload our form
    Unload mfrmAddIn
    
    'point it to nothing
    Set mfrmAddIn = Nothing

End Sub

'Shows the form on startup when needed
Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)

    'get loading settings
    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
    
        'set this to display the form on connect
        Me.Show
        
    End If
    
End Sub

'This event fires when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    
    'When user clicks our button on the toolbar
    Me.Show
    
End Sub

'Adds our add-in button to the Standard toolbar on the VB IDE
Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl

    Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
  
    'handle any errors
    On Error GoTo AddToAddInCommandBarErr
    
    'make sure the standard toolbar is visible
    VBInstance.CommandBars(2).Visible = True
     
    'add the button to the Standard toolbar
    Set cbMenuCommandBar = VBInstance.CommandBars(2).Controls. _
                           Add(1, , , VBInstance.CommandBars(2).Controls.Count)
                           
    'set the caption from string table
    cbMenuCommandBar.Caption = LoadResString(200)
    
    'copy the icon from our resource to the clipboard
    Clipboard.SetData LoadResPicture(1000, 0)
    
    'set the icon on the button
    cbMenuCommandBar.PasteFace
    
    'connect to menucommandbar
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
'handle errors
AddToAddInCommandBarErr:

End Function

