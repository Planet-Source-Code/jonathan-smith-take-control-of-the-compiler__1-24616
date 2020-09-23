VERSION 5.00
Begin VB.Form frmControlPanel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " VB Compilation Controller"
   ClientHeight    =   6780
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   6585
   Icon            =   "ControlPanel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSelectModulesToIntercept 
      Caption         =   "Select Modules to Intercept:"
      Height          =   315
      Left            =   1620
      TabIndex        =   11
      Top             =   3840
      Width           =   2385
   End
   Begin VB.CommandButton cbFinishCompile 
      Caption         =   "Finish Compile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4980
      TabIndex        =   10
      Top             =   90
      Width           =   1515
   End
   Begin VB.CommandButton cbSkipLink 
      Caption         =   "Skip to Linking"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4980
      TabIndex        =   9
      Top             =   480
      Width           =   1515
   End
   Begin VB.CheckBox chkGenerateListing 
      Caption         =   "Generate Listing"
      Height          =   315
      Left            =   1620
      TabIndex        =   8
      Top             =   3600
      Width           =   1905
   End
   Begin VB.TextBox tbApplication 
      Height          =   765
      Left            =   1530
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   120
      Width           =   3345
   End
   Begin VB.CommandButton cbNextModule 
      Caption         =   "Next Module"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4980
      TabIndex        =   5
      Top             =   870
      Width           =   1515
   End
   Begin VB.ListBox lstProjectMembers 
      Height          =   2490
      IntegralHeight  =   0   'False
      Left            =   1560
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   4200
      Width           =   4935
   End
   Begin VB.TextBox tbCommandLine 
      Height          =   2205
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1260
      Width           =   4935
   End
   Begin VB.CommandButton cbSelectAll 
      Caption         =   "Select All"
      Height          =   345
      Left            =   4260
      TabIndex        =   1
      Top             =   3720
      Width           =   1035
   End
   Begin VB.CommandButton cbSelectNone 
      Caption         =   "Select None"
      Height          =   345
      Left            =   5340
      TabIndex        =   0
      Top             =   3720
      Width           =   1155
   End
   Begin VB.Label lblDisplayCurrentModule 
      Caption         =   "[current module]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1620
      TabIndex        =   13
      Top             =   960
      Width           =   3225
   End
   Begin VB.Label lblCurrentModule 
      Alignment       =   1  'Right Justify
      Caption         =   "Current Module:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   1365
   End
   Begin VB.Label lblApplication 
      Alignment       =   1  'Right Justify
      Caption         =   "Application:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   60
      TabIndex        =   7
      Top             =   180
      Width           =   1425
   End
   Begin VB.Label lblCommandLine 
      Alignment       =   1  'Right Justify
      Caption         =   "Command Line:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   90
      TabIndex        =   3
      Top             =   1290
      Width           =   1365
   End
End
Attribute VB_Name = "frmControlPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const DIALOG_CAPTION = " VB Compilation Controller"

Public mIDE As vbide.VBE

Dim mbCancel As Boolean

Public msUserAction As String
Public msUserModules As String

Private masFileNames() As String

Public Sub Activate(sApplication As String, sCommandLine As String, bRefreshModuleList As Boolean)
    chkGenerateListing.Value = vbUnchecked
    tbApplication.Text = sApplication
    tbCommandLine.Text = sCommandLine
    If bRefreshModuleList Then 'beginning of compile
        mbCancel = False
        Me.Caption = DIALOG_CAPTION '& "  " & App.Major & "." & App.Minor & "." & App.Revision
        RefreshList
    End If
    SetModuleName   'determine and display current module
    Me.Show vbModal
    If mbCancel Then
        'do not change command line
    Else
         sApplication = tbApplication.Text
         sCommandLine = tbCommandLine.Text
    End If
End Sub

Private Sub RefreshList()
    Dim aCComponent As VBComponent, idxFileName As Long
On Error GoTo EH
    lstProjectMembers.Clear
    ReDim masFileNames(0 To mIDE.ActiveVBProject.VBComponents.Count)
    idxFileName = 1
    For Each aCComponent In mIDE.ActiveVBProject.VBComponents
        Select Case aCComponent.Type
            Case vbext_ct_ClassModule, vbext_ct_StdModule, vbext_ct_VBForm, vbext_ct_MSForm
                lstProjectMembers.AddItem aCComponent.Name
                masFileNames(idxFileName) = aCComponent.FileNames(1)
                idxFileName = idxFileName + 1
            Case Else
                'no code to check
        End Select
    Next
    Exit Sub
EH:
    MsgBox "Unexpected error refreshing module list: " & Err.Description
End Sub

Private Sub SetGenerateListing()
    Dim sCommandLine As String, posObjectFileSwitch As Long, sObjectFile As String
    sCommandLine = tbCommandLine.Text
    If chkGenerateListing.Value = vbChecked Then
        posObjectFileSwitch = InStr(sCommandLine, "-Fo")
        If posObjectFileSwitch > 1 And InStr(sCommandLine, "-FA") = 0 Then
            sObjectFile = Mid$(sCommandLine, posObjectFileSwitch + 3, InStr(posObjectFileSwitch + 4, sCommandLine, Chr(34)) - posObjectFileSwitch - 2)
            tbCommandLine.Text = Left$(sCommandLine, posObjectFileSwitch - 1) & "-FAs -Fa" & Substitute(sObjectFile, ".obj", ".lst") & " " & Right$(sCommandLine, Len(sCommandLine) - posObjectFileSwitch + 1)
        End If
    Else
        posObjectFileSwitch = InStr(sCommandLine, "-Fo")
        If posObjectFileSwitch > 1 And InStr(sCommandLine, "-FA") > 0 Then
            sObjectFile = Mid$(sCommandLine, posObjectFileSwitch + 3, InStr(posObjectFileSwitch + 4, sCommandLine, """") - 3)
            tbCommandLine.Text = Left$(sCommandLine, InStr(sCommandLine, "-FA") - 1) & Right$(sCommandLine, Len(sCommandLine) - posObjectFileSwitch + 1)
        End If
    End If
    DoEvents
End Sub

Private Sub SetModuleName()
    Dim sCommandLine As String, posObjectFileSwitch As Long, sObjectFilePath As String
    sCommandLine = tbCommandLine.Text
    posObjectFileSwitch = InStr(sCommandLine, "-f ")
    If posObjectFileSwitch > 1 Then
        sObjectFilePath = Mid$(sCommandLine, posObjectFileSwitch + 4, InStr(posObjectFileSwitch + 4, sCommandLine, Chr(34)) - posObjectFileSwitch - 4)
        lblDisplayCurrentModule.Caption = DetermineModuleNameFromPath(sObjectFilePath)
    Else
        lblDisplayCurrentModule.Caption = ""
    End If
End Sub

Private Function DetermineModuleNameFromPath(sFilePath As String) As String
    Dim aCComponent As VBComponent, idxFileName As Long
On Error Resume Next    'non-critical section
    ReDim masFileNames(0 To mIDE.ActiveVBProject.VBComponents.Count)
    idxFileName = 1
    For Each aCComponent In mIDE.ActiveVBProject.VBComponents
        For idxFileName = 1 To aCComponent.FileCount
            If aCComponent.FileNames(idxFileName) = sFilePath Then
                DetermineModuleNameFromPath = aCComponent.Name
                Exit Function
            End If
        Next
    Next
    DetermineModuleNameFromPath = ""    'if not found, then the empty string
End Function

Private Function Substitute(sSource As String, sReplace As String, sWith As String) As String
    Dim posInstr As Long
    Substitute = sSource
    posInstr = InStr(1, Substitute, sReplace, vbTextCompare)
    Do While posInstr > 0
        Substitute = Left$(Substitute, posInstr - 1) & sWith & Right(Substitute, Len(Substitute) - (posInstr + Len(sReplace) - 1))
        posInstr = InStr(posInstr + Len(sWith), Substitute, sReplace)
    Loop
End Function

'EVENT METHODS

Private Sub Form_Load()
    Set mIDE = modAddIn.theConnection.theIDE
    ReDim masFileNames(0 To 0)
    mbCancel = False
    lblDisplayCurrentModule.Caption = ""
End Sub

Private Sub cbSelectAll_Click()
    Dim idxListMember As Long
    For idxListMember = 0 To lstProjectMembers.ListCount - 1
        lstProjectMembers.Selected(idxListMember) = True
    Next
End Sub

Private Sub cbSelectNone_Click()
    Dim idxListMember As Long
    For idxListMember = 0 To lstProjectMembers.ListCount - 1
        lstProjectMembers.Selected(idxListMember) = False
    Next
End Sub

Private Function ListSelectedModules() As String
    Dim ctAllModules As Long, idxModule As Long, sModuleList As String
On Error Resume Next
    ctAllModules = lstProjectMembers.ListCount
    For idxModule = 1 To ctAllModules
        If lstProjectMembers.Selected(idxModule - 1) Then
            sModuleList = sModuleList & masFileNames(idxModule) & Chr(&HFF)
        End If
    Next
    If sModuleList = "" Then sModuleList = Chr(&HFF)
    ListSelectedModules = sModuleList
End Function

Private Sub cbFinishCompile_Click()
    msUserAction = "Finish Compile"
    Me.Hide
End Sub

Private Sub cbSkipLink_Click()
    msUserAction = "Skip to Link"
    Me.Hide
End Sub

Private Sub cbNextModule_Click()
    msUserAction = "Next Module"
    If chkSelectModulesToIntercept.Value = vbChecked Then 'send module list
        msUserModules = ListSelectedModules
    Else
        msUserAction = ""
    End If
    Me.Hide
End Sub

Private Sub chkGenerateListing_Click()
    SetGenerateListing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    msUserAction = "Cancel"
    mbCancel = True
End Sub

