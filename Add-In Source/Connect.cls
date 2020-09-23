VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Compile Controller"
Option Explicit

Implements IDTExtensibility

Const AddInDESCRIPTION = "Compile Controller"
Const ExtraInstanceDESCRIPTION = "Extra Compile Controller"

Private mtheIDE As vbide.VBE
Private mtheMenuHandler As clsMenuHandler

Friend Property Get theIDE() As vbide.VBE
    Set theIDE = mtheIDE
End Property

Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, ByVal ConnectMode As vbext_ConnectMode, ByVal AddInInst As vbide.AddIn, custom() As Variant)
    
    Dim CurrentAddIn As AddIn
    
    On Error GoTo LocalError

    Set mtheIDE = VBInst

    'if the add-in is already loaded we need to terminate this instance of it
    'because multiple instances of the add-in are not supported
    For Each CurrentAddIn In mtheIDE.Addins
        If CurrentAddIn.Description = AddInDESCRIPTION Then 'Add-In is already loaded / terminate
            MsgBox "The compile controller is already loaded."
            AddInInst.Description = ExtraInstanceDESCRIPTION
            Exit Sub
        End If
    Next
    AddInInst.Description = AddInDESCRIPTION
    
    'establish the current environment
    Set modAddIn.theConnection = Me

    'initialize menus if necessary
    If mtheMenuHandler Is Nothing Then
        Set mtheMenuHandler = New clsMenuHandler
        mtheMenuHandler.SetupMenus
    End If

    InitializeAddIn                     'application-specific initializations

'    Select Case ConnectMode
'        Case vbext_cm_Startup          'Add-in was started before the initial Open Project dialog box was shown.
'        Case vbext_cm_AfterStartup     'Add-in was started after the initial Open Project dialog box was shown.
'        Case vbext_cm_External         'Add-in was started externally by another program or component.
'    End Select
  
    Exit Sub
    
LocalError:
    CreateError "Unexpected error while connecting add-in to the VB IDE: " & Err.Description
    
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As vbext_DisconnectMode, custom() As Variant)
    
    On Error Resume Next
    
    mtheMenuHandler.RemoveMenus

    modAddIn.DisconnectAddIn

End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
'
End Sub

Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)
'
End Sub

