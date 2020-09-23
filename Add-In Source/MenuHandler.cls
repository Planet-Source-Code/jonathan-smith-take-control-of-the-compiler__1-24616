VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Dim mIDE                        As vbide.VBE
Dim mcolMenus                   As Collection               'store all menus for easy deletion
Public WithEvents MenuHandler   As CommandBarEvents         'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1

'The caption of our menu choice needs to be toggled so we expose it at the module level.
'If the caption did not change we could have made a local in "SetupMenus()":

Dim menuControlPanel As Office.CommandBarControl  'command bar object

Private Sub Class_Initialize()
    Set mIDE = modAddIn.theConnection.theIDE
End Sub

Friend Sub SetupMenus()

    Dim cbFileMenu As Office.CommandBarControl, cbMakeMenu As Office.CommandBarControl
    Dim ctFileCommandBars As Long, idxCommandBar As Long

On Error GoTo LocalError

    Set mcolMenus = New Collection

    'see if we can find the File menu
    Set cbFileMenu = mIDE.CommandBars(1).Controls("File")    'menu bar is always first command bar
    If cbFileMenu Is Nothing Then
        CreateError "Unable to find the 'File' menu in the VB IDE."
        Exit Sub
    Else 'look for the "Make ..." command
         '(since the caption of this command bar changes dynamically we cannot access it by name
         ' conveniently the way we do with the "File" command bar)
        ctFileCommandBars = cbFileMenu.Controls.Count
        idxCommandBar = 0
        For Each cbMakeMenu In cbFileMenu.Controls
            idxCommandBar = idxCommandBar + 1
            If Left$(cbMakeMenu.Caption, 4) = "Make" Then Exit For 'we found if
            If idxCommandBar = ctFileCommandBars Then 'it was not found
                CreateError "Unable to find the 'Make ...' menu in the VB IDE."
                Exit Sub
            End If
        Next
    End If

    Set menuControlPanel = cbFileMenu.Controls.Add(, , , idxCommandBar) 'add the menu after the Make ... menu
    mcolMenus.Add menuControlPanel                                      'store reference to menu for easy deletion
    
    SetHookMenuCaption False                              'set the caption

    'Note that here we are adding a menu choice that does an invisible internal action.
    'Often you will rather want to have the menu choice start a control panel.  In that case
    'you would probably want to set the menu caption like this:
    'menuControlPanel.Caption = ADDIN_NAME               'set the caption

    Set Me.MenuHandler = mIDE.Events.CommandBarEvents(menuControlPanel)

    Exit Sub
    
LocalError:
    CreateError "Unexpected error while attempting to add " & ADDIN_NAME & " menu to VB IDE: " & Err.Description & " " & Erl
End Sub

Friend Sub SetHookMenuCaption(ByVal bOnOff As Boolean)
    If menuControlPanel Is Nothing Then Exit Sub        'extra protection
    If bOnOff Then
        menuControlPanel.Caption = "Unhook Compilation"
    Else
        menuControlPanel.Caption = "Hook Compilation"
    End If
End Sub

Friend Sub RemoveMenus()
    For Each CurrentMenu In mcolMenus
        CurrentMenu.Delete
    Next
End Sub

'Event occurs when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Select Case CommandBarControl.Caption
        Case "Hook Compilation", "Unhook Compilation"
            SetHookMenuCaption ToggleCreateProcessHook
        Case Else
            'not a handled menu
    End Select
End Sub


