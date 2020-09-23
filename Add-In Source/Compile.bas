Attribute VB_Name = "modCompile"
Option Explicit

'#Const bDEBUG = True

Public Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
End Type

Public Type STARTUPINFO
        cb As Long
        lpReserved As Long      'String
        lpDesktop As Long       'String
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

Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

'Module signatures from Randy Kath's article on PE file format
Public Const IMAGE_DOS_SIGNATURE = &H5A4D       'MZ    short
Public Const IMAGE_OS2_SIGNATURE = &H454E       'NE    short
Public Const IMAGE_OS2_SIGNATURE_LE = &H454C    'LE    short
Public Const IMAGE_NT_SIGNATURE = &H4550        '--PE  long

'Memory-Related public constants from WinNT.H
Public Const PAGE_NOACCESS = &H1
Public Const PAGE_READONLY = &H2
Public Const PAGE_READWRITE = &H4
Public Const PAGE_WRITECOPY = &H8
Public Const PAGE_EXECUTE = &H10
Public Const PAGE_EXECUTE_READ = &H20
Public Const PAGE_EXECUTE_READWRITE = &H40
Public Const PAGE_EXECUTE_WRITECOPY = &H80
Public Const PAGE_GUARD = &H100
Public Const PAGE_NOCACHE = &H200
Public Const PAGE_WRITECOMBINE = &H400
Public Const MEM_COMMIT = &H1000
Public Const MEM_RESERVE = &H2000
Public Const MEM_DECOMMIT = &H4000
Public Const MEM_RELEASE = &H8000
Public Const MEM_FREE = &H10000
Public Const MEM_PRIVATE = &H20000
Public Const MEM_MAPPED = &H40000
Public Const MEM_RESET = &H80000
Public Const MEM_TOP_DOWN = &H100000
Public Const MEM_4MB_PAGES = &H80000000
Public Const SEC_FILE = &H800000
Public Const SEC_IMAGE = &H1000000
Public Const SEC_VLM = &H2000000
Public Const SEC_RESERVE = &H4000000
Public Const SEC_COMMIT = &H8000000
Public Const SEC_NOCACHE = &H10000000
Public Const MEM_IMAGE = SEC_IMAGE

Declare Sub DebugBreak Lib "Kernel32" ()
Declare Function GetModuleHandle Lib "Kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Declare Function GetProcAddress Lib "Kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Declare Function GetLastError Lib "Kernel32" () As Long
Declare Function CreateProcess Lib "Kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Declare Function VirtualProtect Lib "Kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Declare Function lstrlen Lib "Kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Declare Function lenCString Lib "Kernel32" Alias "lstrlenA" (lpString As Long) As Long
Declare Function lstrcpyn Lib "Kernel32" Alias "lstrcpynA" (ByVal lpStringDestination As String, ByVal lpStringSource As String, ByVal lngMaxLength As Long) As Long
Declare Function CopyCString Lib "Kernel32" Alias "lstrcpynA" (ByVal lpStringDestination As String, lpStringSource As Long, ByVal lngMaxLength As Long) As Long

Dim mlpEntryPoint_CreateProcess As Long
Dim mlpFilterLocation As Long

Dim mbCreateProcessHooked As Boolean
Dim mbCompileInProgress As Boolean

'Returns state of hook (True = hooked)
Public Function ToggleCreateProcessHook() As Boolean
    If mbCompileInProgress Then
        frmControlPanel.Show
    Else
        If mbCreateProcessHooked Then
            If UnhookCreateProcess Then
                mbCreateProcessHooked = False
            End If
        Else
            If HookCreateProcess Then
                mbCreateProcessHooked = True
            End If
        End If
        ToggleCreateProcessHook = mbCreateProcessHooked
    End If
End Function

Public Function HookCreateProcess() As Boolean

    Dim sHookError As String, sModuleNameVBA As String
    
    If mbCreateProcessHooked Then Exit Function  'we are already hooked
    mlpFilterLocation = ReturnProcedureAddress(AddressOf CreateProcessFilter)
    
    'Determine name of VBA module (usually "VBA5.DLL" or "VBA6.DLL")
    sModuleNameVBA = GetModuleNameVBA
    If sModuleNameVBA = "" Then
        MsgBox "Unable to determine name of VBA module."
        HookCreateProcess = mbCreateProcessHooked
        Exit Function
    End If

    'Ready to set hook
    If HookDLLImport(sModuleNameVBA, "kernel32", "CreateProcessA", mlpFilterLocation, mlpEntryPoint_CreateProcess, sHookError) Then
        MsgBox "Compiler is hooked."
        HookCreateProcess = True
    Else
        MsgBox "Failed to hook compiler: " & sHookError
        HookCreateProcess = mbCreateProcessHooked
    End If

End Function

'This is my not-very-scientific method for determining which version of the
'VBA DLL is loaded. Hopefully it will be forward-compatible to VB7.
Private Function GetModuleNameVBA() As String
    Dim idxVersion As Long, sModuleName As String
    For idxVersion = 5 To 9
        sModuleName = "VBA" & idxVersion & ".DLL"
        If GetModuleHandle(sModuleName) > 0 Then
            GetModuleNameVBA = sModuleName
            Exit Function
        End If
    Next
End Function

Public Function UnhookCreateProcess(Optional bAnnounce As Boolean = True) As Boolean
    Dim lpFilterLocation As Long, sHookError As String, sModuleNameVBA As String
    If Not mbCreateProcessHooked Then Exit Function
    sModuleNameVBA = GetModuleNameVBA
    If HookDLLImport(sModuleNameVBA, "", mlpFilterLocation, mlpEntryPoint_CreateProcess, lpFilterLocation, sHookError) Then
        If bAnnounce Then MsgBox "Compilation has been unhooked."
        UnhookCreateProcess = True
    Else
        If bAnnounce Then MsgBox "Failed to unhook CreateProcess: " & sHookError
        UnhookCreateProcess = mbCreateProcessHooked
    End If
End Function

'Hooking DLL Calls
'by John Chamberlain
'
'You can use the logic in this function to hook imports in most DLLs and EXEs
'(not just in VB). It will work for most normal Win32 modules. If you use this
'function in your own code please credit its author (me!) and include this
'descriptive header so future users will know what it does.
'
'The call addresses for all implicitly linked DLLs are located in a table
'called the "Import Address Table (IAT)" (or the "Thunk" table). This table is
'generally located at module offset 0x1000 in both DLLs and EXEs and contains
'the addresses of all imported calls in a continuous list with exports from
'different modules separated by NULL (0x 0000 0000). When each DLL is loaded
'the operating system's loader patches this table with the correct addresses.
'In most PE file types an offset to the entry point (which is just past the
'IAT) is located at offset 0xDC from the PE file header which has a signature
'of 0x00004550 (="PE"). Thus the function finds the end of the IAT by scanning
'for this signature and locating the offset.
'
'This function hooks a DLL call by first getting the proc address for the
'specified call and then scanning the IAT for the address. If it is found
'the function substitutes the hook address into the table and returns the
'original address to the caller by reference (in case the caller wants to
'restore the IAT entry to its original state at a later time). If the
'return value was false then the hook could not be set and the reason will
'be returned by reference in the string sError.
'
'When you want to restore the hooked address pass the hook address as
'vCallNameOrAddress and the original address (to be restored) as lpHook.
'The function will find the hooked address in the table and replace it with
'the original address (see UnhookCreateProcess for an example).
'
Private Function HookDLLImport(sImportingModuleName As String, sExportingModuleName As String, vCallNameOrAddress As Variant, lpHook As Long, ByRef lpOriginalAddress As Long, ByRef sError As String) As Boolean
    
    Dim sCallName As String
    Dim lpImportingModuleHandle As Long, lpExportingModuleHandle As Long, lpProcAddress As Long
    Dim vectorIAT As Long, lenIAT As Long, lpEndIAT As Long, lpIATCallAddress As Long
    Dim lpflOldProtect As Long, lpflOldProtect2 As Long
    Dim lpPEHeader As Long
    
    On Error GoTo EH

    'Validate the hook
    If lpHook = 0 Then sError = "Hook is null.": Exit Function

    'Get handle (address) of importing module
    lpImportingModuleHandle = GetModuleHandle(sImportingModuleName)
    If lpImportingModuleHandle = 0 Then sError = "Unable to obtain importing module handle for """ & sImportingModuleName & """.": Exit Function

    'Get the proc address of the IAT entry to be changed
    If VarType(vCallNameOrAddress) = vbString Then
    
        sCallName = CStr(vCallNameOrAddress)    'user is hooking an import
    
        'Get handle (address) of exporting module
        lpExportingModuleHandle = GetModuleHandle(sExportingModuleName)
        If lpExportingModuleHandle = 0 Then sError = "Unable to obtain exporting module handle for """ & sExportingModuleName & """.": Exit Function
    
        'Get address of call
        lpProcAddress = GetProcAddress(lpExportingModuleHandle, sCallName)
        If lpProcAddress = 0 Then sError = "Unable to obtain proc address for """ & sCallName & """.": Exit Function
    
    Else
        lpProcAddress = CLng(vCallNameOrAddress) 'user is restoring a hooked import
    End If

    'Beginning of the IAT is located at offset 0x1000 in most PE modules
    vectorIAT = lpImportingModuleHandle + &H1000

    'Scan module to find PE header by looking for header signature
    lpPEHeader = lpImportingModuleHandle
    Do
        If lpPEHeader > vectorIAT Then  'this is not a PE module
            sError = "Module """ & sImportingModuleName & """ is not a PE module."
            Exit Function
        Else
            If Deref(lpPEHeader) = IMAGE_NT_SIGNATURE Then  'we have located the module's PE header
                Exit Do
            Else
                lpPEHeader = lpPEHeader + 1 'keep searching
            End If
        End If
    Loop
    
    'Determine and validate length of the IAT. The length is at offset 0xDC in the PE header.
    lenIAT = Deref(lpPEHeader + &HDC)
    If lenIAT = 0 Or lenIAT > &HFFFFF Then 'its too big or too small to be valid
        sError = "The calculated length of the Import Address Table in """ & sImportingModuleName & """ is not valid: " & lenIAT
        Exit Function
    End If

    'Scan Import Address Table for proc address
    lpEndIAT = lpImportingModuleHandle + &H1000 + lenIAT
    Do
        If vectorIAT > lpEndIAT Then 'we have reached the end of the table
            sError = "Proc address " & Hex(lpProcAddress) & " not found in Import Address Table of """ & sImportingModuleName & """."
            Exit Function
        Else
            lpIATCallAddress = Deref(vectorIAT)
            If lpIATCallAddress = lpProcAddress Then  'we have found the entry
                Exit Do
            Else
                vectorIAT = vectorIAT + 4   'try next entry in table
            End If
        End If
    Loop
    
    'Substitute hook for existing call address and return existing address by ref
    'We must make this memory writable to make the entry in the IAT
    If VirtualProtect(ByVal vectorIAT, 4, PAGE_EXECUTE_READWRITE, lpflOldProtect) = 0 Then
        sError = "Unable to change IAT memory to execute/read/write."
        Exit Function
    Else
        lpOriginalAddress = Deref(vectorIAT)    'save original address
        CopyMemory ByVal vectorIAT, lpHook, 4   'set the hook
        VirtualProtect ByVal vectorIAT, 4, lpflOldProtect, lpflOldProtect2  'restore memory protection
    End If

    HookDLLImport = True 'mission accomplished

    Exit Function
    
EH:
    sError = "Unexpected error: " & Err.Description

End Function

Function Deref(lngPointer As Long) As Long  'Equivalent of *lngPointer (returns the value pointed to)
    Dim lngValueAtPointer As Long
    CopyMemory lngValueAtPointer, ByVal lngPointer, 4
    Deref = lngValueAtPointer
End Function

Private Function ReturnProcedureAddress(lngAddress As Long) As Long
    ReturnProcedureAddress = lngAddress
End Function

'Declaration Re-Typing for VB

'ByVal lpApplicationName As String          =>  Long (pointer to lpstr)
'ByVal lpCommandLine As String              =>  Long (pointer to lpstr)
'lpProcessAttributes As SECURITY_ATTRIBUTES OK as is
'lpThreadAttributes As SECURITY_ATTRIBUTES  OK as is
'ByVal bInheritHandles As Long              OK as is
'ByVal dwCreationFlags As Long              OK as is
'lpEnvironment As Long                      OK as is
'ByVal lpCurrentDirectory As String         =>  Long (pointer to lpstr)
'lpStartupInfo As STARTUPINFO                   OK as is
'lpProcessInformation As PROCESS_INFORMATION    OK as is

Public Function CreateProcessFilter(lpApplicationName As Long, lpCommandLine As Long, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
    
    Static sInterceptMode As String
    Static bCheckModuleList As Boolean
    Dim bRefreshModuleList As Boolean
    Dim asModuleList() As String
    Dim idxModuleList As Long
    Dim sApplicationName As String, sCommandLine As String, sCurrentDirectory As String, lngCreateProcessReturnValue As Long, lngErrorCode As Long

    sApplicationName = CStringToVBString(lpApplicationName)
    sCommandLine = CStringToVBString(lpCommandLine)
    sCurrentDirectory = CStringToVBString(lpCurrentDirectory)

'#If bDEBUG Then
'    Static sDebugString As String
'    sDebugString = sDebugString & sCommandLine & vbNewLine
'    sDebugString = sDebugString & "intercept mode: " & sInterceptMode & vbNewLine
'    sDebugString = sDebugString & "bcheckmodlist: " & IIf(bCheckModuleList, "true", "false") & vbNewLine
'    MsgBox sDebugString
'#End If

    If mbCompileInProgress Then
        bRefreshModuleList = False  'the module should be current for this compile
    Else
        sInterceptMode = ""         'initialize UI mode settings
        bCheckModuleList = False
        bRefreshModuleList = True   'at the beginning of a compile cycle must refresh module list
        mbCompileInProgress = True  'the compile has started
    End If
    
    If bCheckModuleList Then
        asModuleList = Split(sInterceptMode, Chr(&HFF))
        For idxModuleList = 0 To UBound(asModuleList) - 1
'#If bDEBUG Then
'    sDebugString = sDebugString & "checking: " & """" & asModuleList(idxModuleList) & """" & vbNewLine
'#End If
            If InStr(sCommandLine, """" & asModuleList(idxModuleList) & """") > 0 Then 'module is there
                sInterceptMode = ShowControlPanel(sApplicationName, sCommandLine, bRefreshModuleList)
            End If
        Next
        If Left(sCommandLine, 4) = "LINK" Then
            sInterceptMode = ShowControlPanel(sApplicationName, sCommandLine, bRefreshModuleList)
        End If
    Else
      Select Case sInterceptMode
        Case "", "Next Module"
            sInterceptMode = ShowControlPanel(sApplicationName, sCommandLine, bRefreshModuleList)
        Case "Skip to Link"
            If Left(sCommandLine, 4) = "LINK" Then
                sInterceptMode = ShowControlPanel(sApplicationName, sCommandLine, bRefreshModuleList)
            Else
                'keep compiling
            End If
        Case "Finish Compile"
            'do nothing - just keep calling create process
        Case Else
      End Select
    End If

    If Right(sInterceptMode, 1) = Chr(&HFF) Then
        bCheckModuleList = True
    Else
        bCheckModuleList = False
    End If
    
'#If bDEBUG Then
'    sDebugString = sDebugString & "Executing: " & sCommandLine & vbNewLine
'    sDebugString = sDebugString & vbNewLine
'#End If
    
    'Check to see if ready to link, and if so, reset everything
    If Left(sCommandLine, 4) = "LINK" Then 'this is the end of the compile
        HideControlPanel 'always hide cp when ready to link
        sInterceptMode = "" 'reset intercept mode
        mbCompileInProgress = False
'#If bDEBUG Then
'        Dim hFile As Integer
'        hFile = FreeFile
'        Open "c:\temp\debugcc.tmp" For Output As hFile
'        Write #hFile, sDebugString
'        Close hFile
'#End If
    End If

    lngCreateProcessReturnValue = CreateProcess(sApplicationName, sCommandLine, lpProcessAttributes, lpThreadAttributes, bInheritHandles, dwCreationFlags, lpEnvironment, sCurrentDirectory, lpStartupInfo, lpProcessInformation)
    
    If lngCreateProcessReturnValue = 0 Then
        lngErrorCode = GetLastError
        Select Case lngErrorCode
            Case Else
                CreateError "Error creating process: " & lngErrorCode
        End Select
    End If
    CreateProcessFilter = lngCreateProcessReturnValue

End Function

Function CStringToVBString(lpCString As Long) As String
    Dim lenString As Long, sBuffer As String, lpBuffer As Long, lngStringPointer As Long, refStringPointer As Long
    If lpCString = 0 Then
        CStringToVBString = vbNullString
    Else
        lenString = lenCString(lpCString)
        sBuffer = String$(lenString + 1, 0) 'buffer has one extra byte for terminator
        lpBuffer = CopyCString(sBuffer, lpCString, lenString + 1)
        CStringToVBString = sBuffer
    End If
End Function


