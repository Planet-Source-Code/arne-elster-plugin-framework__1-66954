VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlugInLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function LoadLibrary Lib "kernel32" _
Alias "LoadLibraryA" ( _
    ByVal lpszLib As String _
) As Long

Private Declare Function FreeLibrary Lib "kernel32" ( _
    ByVal hMod As Long _
) As Long

Private Declare Function GetProcAddress Lib "kernel32" ( _
    ByVal hMod As Long, _
    ByVal lpszFnc As String _
) As Long

Private Declare Function CallWindowProc Lib "user32" _
Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
) As Long

Private Declare Function VirtualAlloc Lib "kernel32" ( _
    ByVal lpAddress As Long, _
    ByVal dwSize As Long, _
    ByVal flAllocType As Long, _
    ByVal flProtect As Long _
) As Long

Private Declare Function VirtualFree Lib "kernel32" ( _
    ByVal lpAddress As Long, _
    ByVal dwSize As Long, _
    ByVal dwFreeType As Long _
) As Long

Private Declare Function VirtualProtect Lib "kernel32" ( _
    ByVal lpAddress As Long, _
    ByVal dwSize As Long, _
    ByVal flNewProtect As Long, _
    lpflOldProtect As Long _
) As Long

Private Declare Sub CpyMem Lib "kernel32" _
Alias "RtlMoveMemory" ( _
    pDst As Any, _
    pSrc As Any, _
    ByVal cBytes As Long _
)

Private Declare Function FindFirstFile Lib "kernel32" _
Alias "FindFirstFileA" ( _
    ByVal lpFileName As String, _
    lpFindFileData As WIN32_FIND_DATA _
) As Long

Private Declare Function FindNextFile Lib "kernel32" _
Alias "FindNextFileA" ( _
    ByVal hFindFile As Long, _
    lpFindFileData As WIN32_FIND_DATA _
) As Long

Private Declare Function GetFileAttributes Lib "kernel32" _
Alias "GetFileAttributesA" ( _
    ByVal lpFileName As String _
) As Long

Private Declare Function FindClose Lib "kernel32" ( _
    ByVal hFindFile As Long _
) As Long

Private Const MAX_PATH                  As Long = 260
Private Const MAXDWORD                  As Long = &HFFFF
Private Const FILE_ATTRIBUTE_ARCHIVE    As Long = &H20
Private Const FILE_ATTRIBUTE_DIRECTORY  As Long = &H10
Private Const FILE_ATTRIBUTE_HIDDEN     As Long = &H2
Private Const FILE_ATTRIBUTE_NORMAL     As Long = &H80
Private Const FILE_ATTRIBUTE_READONLY   As Long = &H1
Private Const FILE_ATTRIBUTE_SYSTEM     As Long = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY  As Long = &H100
Private Const INVALID_HANDLE            As Long = -1

Private Enum VirtualFreeTypes
    MEM_DECOMMIT = &H4000
    MEM_RELEASE = &H8000
End Enum

Private Enum VirtualAllocTypes
    MEM_COMMIT = &H1000
    MEM_RESERVE = &H2000
    MEM_RESET = &H8000
    MEM_LARGE_PAGES = &H20000000
    MEM_PHYSICAL = &H100000
    MEM_WRITE_WATCH = &H200000
End Enum

Private Enum VirtualAllocPageFlags
    PAGE_EXECUTE = &H10
    PAGE_EXECUTE_READ = &H20
    PAGE_EXECUTE_READWRITE = &H40
    PAGE_EXECUTE_WRITECOPY = &H80
    PAGE_NOACCESS = &H1
    PAGE_READONLY = &H2
    PAGE_READWRITE = &H4
    PAGE_WRITECOPY = &H8
    PAGE_GUARD = &H100
    PAGE_NOCACHE = &H200
    PAGE_WRITECOMBINE = &H400
End Enum

Private Type allocated_memory
    address     As Long
    bytes       As Long
End Type

Private Type FILETIME
    dwLowDateTime               As Long
    dwHighDateTime              As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes            As Long
    FTCreationTime              As FILETIME
    FTLastAccessTime            As FILETIME
    FTLastWriteTime             As FILETIME
    nFileSizeHigh               As Long
    nFileSizeLow                As Long
    dwReserved0                 As Long
    dwReserved1                 As Long
    cFileName                   As String * MAX_PATH
    cAlternate                  As String * 14
End Type

Private Type PluginClass
    localfile                   As String
    guid                        As UUID
End Type

Private clsInterface            As olelib.IDispatch
Private uidInterface            As UUID
Private strPlugPath             As String
Private strFilter               As String
Private blnRecursive            As Boolean

Private udtPlugins()            As PluginClass
Private lngPluginCnt            As Long

Public Property Get PluginCount() As Long
    PluginCount = lngPluginCnt
End Property

Public Property Get PluginLocation(ByVal index As Long) As String
    If PluginCount = 0 Then Err.Raise 9
    PluginLocation = udtPlugins(index).localfile
End Property

Public Function CreatePlugin(ByVal index As Long) As olelib.IUnknown
    Dim iunkPlugin  As olelib.IUnknown

    If PluginCount = 0 Then Err.Raise 9

    CoCreateInstance udtPlugins(index).guid, _
                     Nothing, _
                     CLSCTX_INPROC_SERVER, _
                     uidInterface, _
                     iunkPlugin

    Set CreatePlugin = iunkPlugin
End Function

Public Function FindPlugins() As Long
    Dim strExts() As String

    lngPluginCnt = 0
    strExts = Split(strFilter, ";")

    FindFilesAPI strPlugPath, strExts, True

    FindPlugins = lngPluginCnt
End Function

Public Property Get filter() As String
    filter = strFilter
End Property

Public Property Let filter(ByVal strF As String)
    strFilter = strF
End Property

Public Property Get Interface() As olelib.IDispatch
    Set Interface = clsInterface
End Property

Public Property Set Interface(clsIdisp As olelib.IDispatch)
    Set clsInterface = clsIdisp
    uidInterface = IIDfromDispatch(clsInterface)
End Property

Public Property Get PluginPath() As String
    PluginPath = strPlugPath
End Property

Public Property Let PluginPath(ByVal strPath As String)
    strPlugPath = strPath
End Property

Public Property Get RecursiveSearch() As Boolean
    RecursiveSearch = blnRecursive
End Property

Public Property Let RecursiveSearch(ByVal blnVal As Boolean)
    blnRecursive = blnVal
End Property

Private Function DirExists(ByVal DirName As String) As Boolean
    On Error Resume Next
    DirExists = GetAttr(DirName) And vbDirectory
End Function

Private Function AddSlash(ByVal strText As String) As String
    AddSlash = IIf(Right$(strText, 1) = "\", strText, strText & "\")
End Function

Private Sub Class_Initialize()
    If DirExists(AddSlash(App.path) & "plugins") Then
        strPlugPath = AddSlash(App.path) & "plugins"
    Else
        strPlugPath = AddSlash(App.path)
    End If

    blnRecursive = True

    strFilter = "*.dll;*.ocx"
End Sub

Private Sub FindFilesAPI( _
    ByVal path As String, _
    filter() As String, _
    ByVal recursive As Boolean _
)

    Dim hSearch     As Long
    Dim udtFindData As WIN32_FIND_DATA
    Dim lngRet      As Long
    Dim i           As Long
    Dim uid         As UUID
    Dim strPFile    As String

    If Not Right$(path, 1) = "\" Then path = path & "\"

    hSearch = FindFirstFile(path & "*.*", udtFindData)
    If hSearch = INVALID_HANDLE Then Exit Sub

    If Left$(udtFindData.cFileName, 1) <> "." Then
        If (udtFindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
            If recursive Then
                FindFilesAPI path & Trim$(StripNulls(udtFindData.cFileName)), filter, recursive
            End If
        Else
            For i = LBound(filter) To UBound(filter)
                If StripNulls(udtFindData.cFileName) Like filter(i) Then
                    strPFile = path & Trim$(StripNulls(udtFindData.cFileName))

                    If IsValidPlugin(strPFile, uid) Then
                        ReDim Preserve udtPlugins(lngPluginCnt) As PluginClass
                        udtPlugins(lngPluginCnt).localfile = strPFile
                        udtPlugins(lngPluginCnt).guid = uid
                        lngPluginCnt = lngPluginCnt + 1
                    End If
                    Exit For
                End If
            Next
        End If
    End If

    lngRet = 1

    Do
        lngRet = FindNextFile(hSearch, udtFindData)
        If lngRet = 0 Then Exit Do

        If Left$(udtFindData.cFileName, 1) <> "." Then
            If (udtFindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                If recursive Then
                    FindFilesAPI path & Trim$(StripNulls(udtFindData.cFileName)), filter, recursive
                End If
            Else
                For i = LBound(filter) To UBound(filter)
                    If StripNulls(udtFindData.cFileName) Like filter(i) Then
                        strPFile = path & Trim$(StripNulls(udtFindData.cFileName))

                        If IsValidPlugin(strPFile, uid) Then
                            ReDim Preserve udtPlugins(lngPluginCnt) As PluginClass
                            udtPlugins(lngPluginCnt).localfile = strPFile
                            udtPlugins(lngPluginCnt).guid = uid
                            lngPluginCnt = lngPluginCnt + 1
                        End If
                        Exit For
                    End If
                Next
            End If
        End If
    Loop

    FindClose hSearch
End Sub

Private Function StripNulls( _
    OriginalStr As String _
) As String

    If InStr(OriginalStr, Chr(0)) > 0 Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If

    StripNulls = Trim$(OriginalStr)
End Function

Private Function IsValidPlugin( _
    ByVal strFile As String, _
    classguid As UUID _
) As Boolean

    Dim clsTypeLib  As ITypeLib
    Dim clsTypeInfo As ITypeInfo
    Dim clsImplInfo As ITypeInfo
    Dim pTypeAttr   As Long
    Dim udtTypeAttr As TYPEATTR
    Dim udtImplAttr As TYPEATTR
    Dim i           As Long
    Dim j           As Long
    Dim hRefType    As Long

    ' first register the component, then try to load its type library
    If Not RegisterServer(strFile, True) Then Exit Function

    Set clsTypeLib = LoadTypeLibEx(strFile, REGKIND_REGISTER)
    If clsTypeLib Is Nothing Then Exit Function

    ' get all the CoClasses in the server and search for
    ' the interface the plugins have to implement
    For i = 0 To clsTypeLib.GetTypeInfoCount - 1
        If clsTypeLib.GetTypeInfoType(i) = TKIND_COCLASS Then
            Set clsTypeInfo = clsTypeLib.GetTypeInfo(i)

            pTypeAttr = clsTypeInfo.GetTypeAttr
            If pTypeAttr <> 0 Then
                CpyMem udtTypeAttr, ByVal pTypeAttr, Len(udtTypeAttr)
                clsTypeInfo.ReleaseTypeAttr pTypeAttr

                ' Implements of the current class
                For j = 0 To udtTypeAttr.cImplTypes - 1
                    hRefType = clsTypeInfo.GetRefTypeOfImplType(j)
                    Set clsImplInfo = clsTypeInfo.GetRefTypeInfo(hRefType)

                    If Not clsImplInfo Is Nothing Then
                        pTypeAttr = clsImplInfo.GetTypeAttr
                        If pTypeAttr <> 0 Then
                            CpyMem udtImplAttr, ByVal pTypeAttr, Len(udtImplAttr)
                            clsImplInfo.ReleaseTypeAttr pTypeAttr

                            If CompareGUIDs(udtImplAttr.iid, uidInterface) Then
                                ' a class implements the specified interface,
                                ' we found a plugin!
                                classguid = udtTypeAttr.iid
                                IsValidPlugin = True
                                Exit Function
                            End If

                        End If
                    End If

                Next

            End If
        End If
    Next
End Function

Private Function CompareGUIDs( _
    guid1 As UUID, _
    guid2 As UUID _
) As Boolean

    Dim i   As Long

    If guid1.Data1 = guid2.Data1 Then
        If guid1.Data2 = guid2.Data2 Then
            If guid1.Data3 = guid2.Data3 Then
                For i = 0 To 7
                    If guid1.Data4(i) <> guid2.Data4(i) Then
                        Exit Function
                    End If
                Next

                CompareGUIDs = True
            End If
        End If
    End If
End Function

Private Function IIDfromDispatch( _
    clsDisp As olelib.IDispatch _
) As UUID

    Dim pTypeAttr   As Long
    Dim udtTypeAttr As TYPEATTR

    pTypeAttr = clsDisp.GetTypeInfo.GetTypeAttr

    If pTypeAttr = 0 Then Exit Function
    CpyMem udtTypeAttr, ByVal pTypeAttr, Len(udtTypeAttr)

    IIDfromDispatch = udtTypeAttr.iid

    clsDisp.GetTypeInfo.ReleaseTypeAttr pTypeAttr
End Function

Private Function RegisterServer( _
    ByVal strFile As String, _
    ByVal register As Boolean _
) As Boolean

    Dim hLib    As Long
    Dim fpReg   As Long

    hLib = LoadLibrary(strFile)
    If hLib = 0 Then Exit Function

    If register Then
        fpReg = GetProcAddress(hLib, "DllRegisterServer")
    Else
        fpReg = GetProcAddress(hLib, "DllUnregisterServer")
    End If

    If fpReg = 0 Then Exit Function

    CallStd fpReg

    RegisterServer = True
End Function

Private Function AllocMemory( _
    ByVal bytes As Long, _
    Optional ByVal lpAddr As Long = 0, _
    Optional ByVal PageFlags As VirtualAllocPageFlags = PAGE_READWRITE _
) As allocated_memory

    With AllocMemory
        .address = VirtualAlloc(lpAddr, bytes, MEM_COMMIT, PageFlags)
        .bytes = bytes
    End With
End Function

Private Function FreeMemory( _
    udtMem As allocated_memory _
) As Boolean

    VirtualFree udtMem.address, udtMem.bytes, MEM_DECOMMIT

    udtMem.address = 0
    udtMem.bytes = 0
End Function

Private Function CallStd( _
    ByVal fnc As Long, _
    ParamArray Params() As Variant _
) As Long

    Dim udtMem              As allocated_memory
    Dim pASM                As Long
    Dim i                   As Integer

    udtMem = AllocMemory(&HEC00&, , PAGE_EXECUTE_READWRITE)
    If udtMem.address = 0 Then Exit Function
    pASM = udtMem.address

    AddByte pASM, &H58                  ' POP EAX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H50                  ' PUSH EAX

    If UBound(Params) = 0 Then
        If IsArray(Params(0)) Then
            For i = UBound(Params(0)) To 0 Step -1
                AddPush pASM, CLng(Params(0)(i))    ' PUSH dword
            Next
        Else
            For i = UBound(Params) To 0 Step -1
                AddPush pASM, CLng(Params(i))       ' PUSH dword
            Next
        End If
    Else
        For i = UBound(Params) To 0 Step -1
            AddPush pASM, CLng(Params(i))           ' PUSH dword
        Next
    End If

    AddCall pASM, fnc                   ' CALL rel addr
    AddByte pASM, &HC3                  ' RET

    CallStd = CallWindowProc(udtMem.address, _
                             0, 0, 0, 0)

    FreeMemory udtMem
End Function

Private Sub AddPush( _
    pASM As Long, _
    lng As Long _
)

    AddByte pASM, &H68
    AddLong pASM, lng
End Sub

Private Sub AddCall( _
    pASM As Long, _
    addr As Long _
)

    AddByte pASM, &HE8
    AddLong pASM, addr - pASM - 4
End Sub

Private Sub AddLong( _
    pASM As Long, _
    lng As Long _
)

    CpyMem ByVal pASM, lng, 4
    pASM = pASM + 4
End Sub

Private Sub AddByte( _
    pASM As Long, _
    Bt As Byte _
)

    CpyMem ByVal pASM, Bt, 1
    pASM = pASM + 1
End Sub
