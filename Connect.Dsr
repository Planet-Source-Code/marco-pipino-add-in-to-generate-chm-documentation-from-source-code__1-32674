VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   10050
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   10260
   _ExtentX        =   18098
   _ExtentY        =   17727
   _Version        =   393216
   Description     =   $"Connect.dsx":0000
   DisplayName     =   "Doc Generator"
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
Attribute VB_HelpID = 2002
'Author: VB Add-in Wizard and Marco Pipino
Option Explicit

Public FormDisplayed          As Boolean
Attribute FormDisplayed.VB_VarHelpID = 2003
Public VBInstance             As VBIDE.VBE
Attribute VBInstance.VB_VarHelpID = 2004
Dim mcbMenuCommandBar         As Office.CommandBarControl
Dim mfrmAddIn                  As New frmAddIn
Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1
Private icnIcon As IPictureDisp

Sub Hide()
Attribute Hide.VB_HelpID = 2005
    On Error Resume Next
    FormDisplayed = False
    mfrmAddIn.Hide
End Sub

Sub Show()
Attribute Show.VB_HelpID = 2006
    On Error Resume Next
    If Len(VBInstance.ActiveVBProject.FileName) > 0 Then
    
        If mfrmAddIn Is Nothing Then
            Set mfrmAddIn = New frmAddIn
        End If
        
        Set mfrmAddIn.VBInstance = VBInstance
        Set mfrmAddIn.Connect = Me
        FormDisplayed = True
        Load mfrmAddIn
        mfrmAddIn.Show
    Else
        MsgBox ("You must save the project")
    End If
End Sub

'Purpose:this method adds the Add-In to VB
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    'save the vb instance
    Set VBInstance = Application
    
    'this is a good place to set a breakpoint and
    'test various addin objects, properties and methods

    If ConnectMode = ext_cm_External Then
        'Used by the wizard toolbar to start this wizard
        Me.Show
    Else
        Set mcbMenuCommandBar = AddToAddInCommandBar("Documentation Generator")
        
        'sink the event
        Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
    End If
  
    If ConnectMode = ext_cm_AfterStartup Then
        If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
            'set this to display the form on connect
            Me.Show
        End If
    End If
  
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
    
End Sub

'Purpose:this method removes the Add-In from VB
Private Sub AddinInstance_OnDisconnection( _
        ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, _
        custom() As Variant)
        
    On Error Resume Next
    'delete the command bar entry
    mcbMenuCommandBar.Delete
    
    'shut down the Add-In
    If FormDisplayed Then
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
        FormDisplayed = False
    Else
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
    End If
    
    Unload mfrmAddIn
    Set mfrmAddIn = Nothing

End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
        'set this to display the form on connect
        Me.Show
    End If
End Sub

'Purpose:this event fires when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Me.Show
End Sub

'Purpose: Add the Command in the menu with the the hhc bitmap
Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
Attribute AddToAddInCommandBar.VB_HelpID = 2007
    Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
    Dim cbMenu As Object
    
    On Error GoTo AddToAddInCommandBarErr
    
    'see if we can find the Add-Ins menu
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        'not available so we fail
        Exit Function
    End If
    
    'add it to the command bar
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    'set the caption
    cbMenuCommandBar.Caption = sCaption
    
    'Copy the bitmap ...
    Clipboard.SetData LoadPicture(App.Path & "\hhc.bmp")
    'And paste it in the menu
    cbMenuCommandBar.PasteFace
    Set AddToAddInCommandBar = cbMenuCommandBar
    Exit Function
    
AddToAddInCommandBarErr:

End Function

