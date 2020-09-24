Attribute VB_Name = "modGlobals"
'Notes
'################################################
'"a" - means it is an Address
'"o" - means it is a relative Offset
'"Unknown" - self explanatory
'"Flag" - Variable Unknown Property
'"Const" - Constant Unknown Property
'"Address" - Unknown Address
'################################################
Const Signature = &H1F4 ' F4 01 00 00
Const MAX_PATH = 260
Public Const Version = "0.02"


Type VBHeader
Signature               As String * 4  '00h 00d
    'VB5! identifier &quot;VB5!&quot;
    
   Flag1                  As Integer     '04h 04d
     'Seems constant for each machine, _
      event after reformat etc. Changing _
      values do not effect working status of an exe.
  
  LanguageDLL             As String * 14 '06h 06d
    'Language DLL name. _
     0x2A meaning default or null terminated string.
  
  BackupLanguageDLL       As String * 14 '14h 20d
    'Backup Language DLL name. _
     0x7F meaning default or null terminated string. _
     Changing values do not effect working status of an exe.
  
  RuntimeDLLVersion       As Integer     '22h 34d
    'Run-time DLL version
  
  LanguageID              As Long        '24h 36d
  
  BackupLanguageID        As Long        '28h 40d
    'Backup Language ID &#40;only when Language DLL exists&#41;
  
  aSubMain                As Long        '2Ch 44d
    'Address to Sub Main&#40;&#41; code _
     &#40;If 0000 0000 then it's a load form call&#41;
  
  aProjectInfo            As Long        '30h 48d
    
   Flag2                  As Integer     '34h 52d
     'Something that changes with the primary form _
      changes - ie no form/bare form/form with controls. _
      Maybe binary or single hex &#40;nibble&#41; values? _
      Changing values dont effect working status of an exe.
    
   Flag3                  As Integer     '36h 54d
     'Most seem to be 0x30,0x00. _
      Maybe binary or single hex &#40;nibble&#41; values? _
      Changing values dont effect working status of an exe.
      
   Flag4                  As Long        '38h 56d
     'Maybe binary or single hex &#40;nibble&#41; values? _
      Changing values dont effect working status of an exe.
  
  ThreadSpace             As Long        '3Ch 60d
    'Thread space. Changing values do effect working _
     status of an exe to some extent. Some values work, _
     some dont, others make the whole program crash.
      
   Const1                 As Long        '40h 64d
     'Something count?
  
  FormCount               As Integer     '44h 68d
  
  ExternalComponentCount  As Integer     '46h 70d
    'Number of external components &#40;eg. winsock&#41; referenced
    
   Flag5                  As Byte        '48h 72d
     'Maybe this is max. allocatable memory? _
      Low values seem to crash. Changing values do effect _
      working status of an exe.
    
   Flag6                  As Byte        '49h 73d
     'The bigger the number, the longer the wait and more _
      memory used before program starts functioning as _
      normal. Allocated memory? Obviously, the working _
      status of an exe is effected by this flag.
    
   Const2                 As Integer     '4Ah 74d
     'Unknown
  
  aGUITable               As Long        '4Eh 78d
  aExternalComponentTable As Long        '52h 82d
  aProjectDescription     As Long        '56h 86d
  
  oProjectExename         As Long        '5Ah 90d
  oProjectTitle           As Long        '5Eh 94d
  oHelpFile               As Long        '62h 98d
  oProjectName            As Long        '66h 102d
End Type

Private Type tProjectInfo

  Signature As Long                            ' 0x00
  aObjectTable As Long                         ' 0x04
  Null1 As Long                                ' 0x08
  aStartOfCode As Long                         ' 0x0C
  aEndOfCode As Long                           ' 0x10
  Flag1 As Long                                ' 0x14
  ThreadSpace As Long                          ' 0x18
  aVBAExceptionhandler  As Long                ' 0x1C
  aNativeCode As Long                          ' 0x20
  oProjectLocation As Integer                  ' 0x24
  Flag2 As Integer                             ' 0x26
  Flag3 As Integer                             ' 0x28

  OriginalPathName(MAX_PATH * 2) As Byte       ' 0x2A
  NullSpacer As Byte                           ' 0x233
  aExternalTable As Long                       ' 0x234
  ExternalCount As Long                        ' 0x238

' Size 0x23C
End Type

Private Type tObject
    aObjectInfo As Long         ' 0x00
    Const1 As Long              ' 0x04
    Address1 As Long            ' 0x08
    Null1 As Long               ' 0x0C
    Address2 As Long            ' 0x10
    Null2 As Long               ' 0x14
    aObjectName As Long         ' 0x18  NTS
    ProcCount As Long           ' 0x1C events, funcs, subs
    aProcNamesArray As Long     ' 0x20 when non-zero
    Const2 As Long              ' 0x24
    ObjectType As Long          ' 0x28
    Null3 As Long               ' 0x2C
                                ' 0x30  <-- Structure Size
End Type
Public Type tObjectInfo
    Flag1 As Integer       ' 0x00
    ObjectIndex As Integer ' 0x02
    aObjectTable As Long   ' 0x04
    Null1 As Long          ' 0x08
    aSmallRecord   As Long ' 0x0C  when it is a module this value is -1 [better name?]
    Const1 As Long         ' 0x10
    Null2 As Long          ' 0x14
    aObject As Long        ' 0x18
    RunTimeLoaded  As Long ' 0x1C [can someone verify this?]
    NumberOfProcs  As Long ' 0x20
    aProcTable As Long     ' 0x24
    Flag3 As Integer       ' 0x28
    Flag4 As Integer       ' 0x2A
    Flag5 As Long          ' 0x2C
    Flag6 As Integer       ' 0x30
    Flag7 As Integer       ' 0x32
    aConstantPool As Long  ' 0x34
                           ' 0x38 <-- Structure Size
                           'the rest is optional items[OptionalObjectInfo]
End Type
Private Type tObjectTable

    Null1 As Long           ' 0x00
    Address1 As Long        ' 0x04
    Address2 As Long        ' 0x08
    Const1 As Long          ' 0x0C
    Null2 As Long           ' 0x10
    Address3 As Long        ' 0x14
    Flag1 As Long           ' 0x18
    Flag2 As Long           ' 0x1C
    Flag3 As Long           ' 0x20
    Flag4 As Long           ' 0x24
    Const2 As Integer       ' 0x28
    ObjectCount1 As Integer ' 0x2A
    ObjectCount2 As Integer ' 0x2C
    ObjectCount3 As Integer ' 0x2E
    aObject As Long         ' 0x30
    Null3 As Long           ' 0x34
    Null4 As Long           ' 0x38
    Null5 As Long           ' 0x3C
    aProjectName As Long    ' 0x40      NTS
    LangID1  As Long        ' 0x44
    LangID2  As Long        ' 0x48
    Null6  As Long          ' 0x4C
    Const3  As Long         ' 0x50
                            ' 0x54
End Type
Type ExternalTable
   Flag As Long        '0x00
   aExternalLibrary As Long  '0x04
End Type

Type ExternalLibrary
   aLibraryName As Long     '0x00   points to NTS
   aLibraryFunction As Long '0x04   points to NTS
End Type

Private Type tEventLink

    Const1 As Integer        ' 0x00
    CompileType As Byte      ' 0x02
    aEvent As Long           ' 0x03
    PushCmd As Byte          ' 0x07
    PushAddress As Long      ' 0x08
    Const As Byte            ' 0x0C
                             ' 0x0D&lt;-- Structure Size
End Type
Private Type tEventTable
    Null1 As Long                                  ' 0x00
    aControl As Long                               ' 0x04
    aObjectInfo As Long                            ' 0x08
    aQueryInterface As Long                        ' 0x0C
    aAddRef As Long                                ' 0x10
    aRelease As Long                                ' 0x14
    aEventPointer() As Long
    'aEventPointer(aControl.EventCount - 1) As Long ' 0x18
End Type


Private Type tOptionalObjectInfo ' if &#40;&#40;tObject.ObjectType AND &amp;H80&#41;=&amp;H80&#41;

    Const1 As Long          ' 0x00  01 00 00 00
    Address1 As Long        ' 0x04
    Null1 As Long           ' 0x08
    Address2 As Long        ' 0x0C
    Const2 As Long          ' 0x10  01 00 00 00
    Address3 As Long        ' 0x14
    Null2 As Long           ' 0x18
    Address4 As Long        ' 0x1C
    ControlCount As Long    ' 0x20
    aControlArray As Long   ' 0x24
    Const3 As Long          ' 0x28
    Const4 As Long          ' 0x2C
    aEventLinkArray As Long ' 0x30
    Address6 As Long        ' 0x34
    Null3 As Long           ' 0x38
    Flag1 As Long           ' 0x3C usually null
                            ' 0x40 &lt;-- Structure size
End Type
Private Type tEventPointer
    Const1 As Byte      ' 0x00
    Flag1 As Long       ' 0x01
    Const2 As Long      ' 0x05
    Const3 As Byte      ' 0x09
    aEvent As Long      ' 0x0A
                        ' 0x0E &lt;-- Structure Size
End Type

Private Type tCodeInfo
    aObjectInfo As Long     ' 0x00
    Flag1 As Integer        ' 0x04
    Flag2 As Integer        ' 0x06
    CodeLength As Integer   ' 0x08
    Flag3 As Long           ' 0x0A
    Flag4 As Integer        ' 0x0E
    Null1 As Integer        ' 0x10
    Flag5 As Long           ' 0x12
    Flag6 As Integer        ' 0x16
                            ' 0x18  &lt;-- Structure Size
End Type
Public Type tProcedure
  aProcedure1  As Long    ' 0x0
  aProcedure2  As Long    ' 0x4

  aProcedureN  As Long    '
                          ' &#40;ObjectInfo.NumberOfProcs * 4&#41; bytes
End Type
Private Type tControl
    Flag1 As Integer        ' 0x00
    EventCount As Integer   ' 0x02
    Flag2 As Long           ' 0x04
    aGUID As Long           ' 0x08
    Index As Integer        ' 0x0C
    Const1 As Integer       ' 0x0E
    Null1 As Long           ' 0x10
    Null2 As Long           ' 0x14
    aEventTable As Long     ' 0x18
    Flag3 As Byte           ' 0x1C
    Const2 As Byte          ' 0x1D
    Const3 As Integer       ' 0x1E
    aName As Long           ' 0x20
    Index2 As Integer       ' 0x24
    Const1Copy As Integer   ' 0x26
                            ' 0x28  &lt;-- Structure Size
End Type

Private Type tGuiTable
    SectionHeader As Long
    unknown(59) As Byte
    FormSize As Long
    un1 As Long
    aFormPointer As Long
    un2 As Long
    
End Type

'Globals begin
Global gVBHeader As VBHeader
Global gProjectInfo As tProjectInfo
Global gObjectTable As tObjectTable
Global gObject() As tObject
Global gObjectInfo As tObjectInfo
Global gExternalTable As ExternalTable
Global gExternalLibrary As ExternalLibrary
Global gOptionalObjectInfo As tOptionalObjectInfo
Global gEventLink As tEventLink
Global gEventPointer As tEventPointer
Global gControl() As tControl
Global gEventTable() As tEventTable
Global gCodeInfo As tCodeInfo
Global gProcedure()  As Long 'As tProcedure
Global gGuiTable() As tGuiTable
Global gObjectNameArray() As String

Global gSkipCom As Boolean
Global gDumpData As Boolean

Private Type typeControlName
    strParentForm As String
    strControlName As String
End Type
Global gControlNameArray() As typeControlName

Private Type typeExternTable
    Length As Integer
    
End Type
'For Controls
Public Type typeStandardControlSize
    cLeft As Integer
    cTop As Integer
    cWidth As Integer
    cHeight As Integer
End Type

'Picture Header
Public Type typePictureHeader
    un1 As Integer
    un2 As Integer
    un3 As Integer
    un4 As Integer
End Type

'COM Fix Type to fix some COM problems
Private Type typeCOMFIX
    ObjectName As String
    PropertyName As String
    NewType As String
End Type
Global gComFix() As typeCOMFIX
'Used for Memory Map
Public gVBFile As clsFile
Public gMemoryMap As clsMemoryMap

'Variables for .vbp file
Global ProjectExename As String                     ' Project exename. MaxLength: 0x104 (260d)
Global ProjectTitle As String                       ' Project title. MaxLength: 0x28 (40d)
Global HelpFile As String                           ' Helpfile. MaxLength: 0x28 (40d)
Global ProjectName As String                        ' Project name. MaxLength: 0x104 (260d)
Global ProjectDescription As String

'Get File Information File Version Properties
Public Type FILEPROPERTIE
    CompanyName As String
    FileDescription As String
    FileVersion As String
    InternalName As String
    LegalCopyright As String
    OrigionalFileName As String
    ProductName As String
    ProductVersion As String
    LanguageID As String
End Type
Global gFileInfo As FILEPROPERTIE
Declare Function GetFileVersionInfo Lib "Version.dll" Alias _
   "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal _
   dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias _
   "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, _
   lpdwHandle As Long) As Long
Declare Function VerQueryValue Lib "Version.dll" Alias _
   "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, _
   lplpBuffer As Any, puLen As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias _
   "GetSystemDirectoryA" (ByVal Path As String, ByVal cbBytes As _
   Long) As Long
Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Dest As Any, ByVal Source As Long, ByVal Length As Long)
Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" ( _
    ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Public Const LANG_ENGLISH = &H9

Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

Sub PrintReadMe()
    '*****************************
    'Prints the ReadMe of the program
    '*****************************
    On Error Resume Next
    Kill (App.Path & "\readme.txt")
   
    Open App.Path & "\ReadMe.txt" For Output As #1
        Print #1, "-------------------------------"
        Print #1, "Semi VB Decompiler by vbgamer45"
        Print #1, "Open Source"
        Print #1, "Version: " & Version
        Print #1, "-------------------------------"
        Print #1, "Contents"
        Print #1, "1. What's New"
        Print #1, "2. Features"
        Print #1, "3. Questions?"
        Print #1, "4. Bugs"
        Print #1, "5. Contact"
        Print #1, ""
        Print #1, "1. What's New"
        Print #1, "   Version 0.02 Rebuilds the forms"
        Print #1, "   Gets most controls and their properties."
        Print #1, "   Intial Release version 0.01"
        Print #1, ""
        Print #1, "2. Features"
        Print #1, "   Decompiling of native vb6/vb5 exe's"
        Print #1, ""
        Print #1, "3. Questions?"
        Print #1, ""
        Print #1, "4. Bugs"
        Print #1, ""
        Print #1, "5. Contact"
    Close #1
    
End Sub
Public Function sHexStringFromString(ByVal inp As String, Optional Spacing As Boolean = True) As String
Dim hc As String
Dim hs As String
Dim c As Long
While Len(inp)
    
    hc = Hex(Asc(Mid(inp, 1, 1)))
    inp = Mid(inp, 2)
    If Len(hc) = 1 Then hc = "0" & hc
    hs = hs & hc
    c = c + 1
    If Spacing Then
        If c Mod 4 = 0 Then
            hs = hs & "  "
        ElseIf c Mod 2 = 0 Then
            hs = hs & " "
        End If
        
    End If
Wend
sHexStringFromString = hs
End Function
Public Function PadHex(ByVal sHex As String, Optional Pad As Integer = 8) As String
    If Len(sHex) > Pad Then
        PadHex = sHex
    Else
        PadHex = String(Pad - Len(sHex), 48) & sHex
    End If
End Function

Public Function AddChar(Val As String, TheLen As Long, Optional Char As String = "0") As String    'Permet d'ajouter un charactère à une chaine de charactère pour obtenir une certaine longueur.
    AddChar = Right(String(TheLen, Char) & Val, TheLen)
End Function
Public Function ExtString(DataStr As String) As String
    ExtString = Left(DataStr, lstrlen(DataStr))
End Function
Public Function GetUntilNull(FileNum As Variant) As String
    '*****************************
    'Purpose to get a null termintated string
    '*****************************
    Dim aList() As Byte
    Dim k As Byte
    k = 255
    ReDim aList(0)
    Do Until k = 0
        Get FileNum, , k
        ReDim Preserve aList(UBound(aList) + 1)
        aList(UBound(aList)) = k
        'MsgBox k
    Loop
    Dim i As Integer
    Dim Final As String
    For i = 1 To UBound(aList) - 1
        Final = Final & Chr(aList(i))
      
    Next i
    
    GetUntilNull = Final
End Function
Public Function GetUnicodeString(FileNum As Variant, Length As Integer) As String
    '*****************************
    'Purpose to get a unicode string
    '*****************************
    Dim aList() As Byte

    ReDim aList((Length * 2))
    Get FileNum, , aList

    Dim i As Integer
    Dim Final As String
    For i = 1 To UBound(aList) - 1
        If aList(i) <> 0 Then
            Final = Final & Chr(aList(i))
        End If
    Next i
    
    GetUnicodeString = Final
End Function
Public Function FileInfo(Optional ByVal PathWithFilename As String) As FILEPROPERTIE
 ' return file-properties of given file  (EXE , DLL , OCX)

Static BACKUP As FILEPROPERTIE   ' backup info for next call without filename
If Len(PathWithFilename) = 0 Then
    FileInfo = BACKUP
    Exit Function
End If

Dim lngBufferlen As Long
Dim lngDummy As Long
Dim lngRc As Long
Dim lngVerPointer As Long
Dim lngHexNumber As Long
Dim bytBuffer() As Byte
Dim bytBuff(255) As Byte
Dim strBuffer As String
Dim strLangCharset As String
Dim strVersionInfo(7) As String
Dim strTemp As String
Dim intTemp As Integer
       
' size
lngBufferlen = GetFileVersionInfoSize(PathWithFilename, lngDummy)
If lngBufferlen > 0 Then
   ReDim bytBuffer(lngBufferlen)
   lngRc = GetFileVersionInfo(PathWithFilename, 0&, lngBufferlen, bytBuffer(0))
   If lngRc <> 0 Then
      lngRc = VerQueryValue(bytBuffer(0), "\VarFileInfo\Translation", _
               lngVerPointer, lngBufferlen)
      If lngRc <> 0 Then
         'lngVerPointer is a pointer to four 4 bytes of Hex number,
         'first two bytes are language id, and last two bytes are code
         'page. However, strLangCharset needs a  string of
         '4 hex digits, the first two characters correspond to the
         'language id and last two the last two character correspond
         'to the code page id.
         MoveMemory bytBuff(0), lngVerPointer, lngBufferlen
         lngHexNumber = bytBuff(2) + bytBuff(3) * &H100 + _
                bytBuff(0) * &H10000 + bytBuff(1) * &H1000000
         strLangCharset = Hex(lngHexNumber)
         'now we change the order of the language id and code page
         'and convert it into a string representation.
         'For example, it may look like 040904E4
         'Or to pull it all apart:
         '04------        = SUBLANG_ENGLISH_USA
         '--09----        = LANG_ENGLISH
         ' ----04E4 = 1252 = Codepage for Windows:Multilingual
         'Do While Len(strLangCharset) < 8
         '    strLangCharset = "0" & strLangCharset
         'Loop
         If Mid(strLangCharset, 2, 2) = LANG_ENGLISH Then
         strLangCharset2 = "English (US)"

         
         End If

         Do While Len(strLangCharset) < 8
             strLangCharset = "0" & strLangCharset
         Loop
         
         ' assign propertienames
         strVersionInfo(0) = "CompanyName"
         strVersionInfo(1) = "FileDescription"
         strVersionInfo(2) = "FileVersion"
         strVersionInfo(3) = "InternalName"
         strVersionInfo(4) = "LegalCopyright"
         strVersionInfo(5) = "OriginalFileName"
         strVersionInfo(6) = "ProductName"
         strVersionInfo(7) = "ProductVersion"
         ' loop and get fileproperties
         For intTemp = 0 To 7
            strBuffer = String$(255, 0)
            strTemp = "\StringFileInfo\" & strLangCharset _
               & "\" & strVersionInfo(intTemp)
            lngRc = VerQueryValue(bytBuffer(0), strTemp, _
                  lngVerPointer, lngBufferlen)
            If lngRc <> 0 Then
               ' get and format data
               lstrcpy strBuffer, lngVerPointer
               strBuffer = Mid$(strBuffer, 1, InStr(strBuffer, Chr(0)) - 1)
               strVersionInfo(intTemp) = strBuffer
             Else
               ' property not found
               strVersionInfo(intTemp) = "?"
            End If
         Next intTemp
      End If
   End If
End If
' assign array to user-defined-type
FileInfo.CompanyName = strVersionInfo(0)
FileInfo.FileDescription = strVersionInfo(1)
FileInfo.FileVersion = strVersionInfo(2)
FileInfo.InternalName = strVersionInfo(3)
FileInfo.LegalCopyright = strVersionInfo(4)
FileInfo.OrigionalFileName = strVersionInfo(5)
FileInfo.ProductName = strVersionInfo(6)
FileInfo.ProductVersion = strVersionInfo(7)
FileInfo.LanguageID = strLangCharset2
BACKUP = FileInfo
End Function
'*****************************
'The following functions are used for COM
'*****************************
Public Function GetBoolean(FileNum As Variant) As Boolean
        Dim k As Boolean
        Get FileNum, , k
        GetBoolean = k
End Function
Public Function GetByte2(FileNum As Variant) As Byte
        Dim k As Byte
        Get FileNum, , k
        GetByte2 = k
End Function
Public Function GetInteger(FileNum As Variant) As Integer
        Dim k As Integer
        Get FileNum, , k
        
        GetInteger = k
End Function
Public Function GetLong(FileNum As Variant) As Long
        Dim k As Long
        Get FileNum, , k
        GetLong = k
End Function
Public Function GetSingle(FileNum As Variant) As Single
        Dim k As Single
        Get FileNum, , k
        GetSingle = k
End Function
Public Function GetString(FileNum As Variant) As String
    'Not used...
        Dim k As String
        Seek FileNum, (Loc(FileNum) + 3)
        Get FileNum, , k
  
        GetString = k
End Function
Public Function GetAllString(FileNum As Variant) As String
    Dim Length As Integer
    Get FileNum, , Length
    
    Dim strText As String
    strText = GetUntilNull(FileNum)
    'MsgBox strText
    If Len(strText) < Length Then
    'get unicode string
   ' MsgBox "unicode"
        If Length < 100 Then
            Seek FileNum, Loc(FileNum) - 2
            'MsgBox "Loc: " & Loc(FileNum)
            strText = GetUnicodeString(FileNum, Length)
            'MsgBox "Loc: " & Loc(FileNum)
            Seek FileNum, Loc(FileNum) + 1
        End If
    End If
    GetAllString = strText
End Function

Sub AddText(strText As String)
    frmMain.txtFinal(frmMain.txtFinal.UBound).Text = frmMain.txtFinal(frmMain.txtFinal.UBound).Text & strText & vbCrLf
End Sub
Sub LoadNewFormHolder(FormName As String)
'Purpose to hold each form's information
    Dim i As Integer
    i = frmMain.txtFinal.UBound + 1
    Load frmMain.txtFinal(i)
    With frmMain.txtFinal(i)
        .Tag = FormName
    
    End With
End Sub

Sub LoadCOMFIX()
'Load the COM Hacks
'Com Hack File Format
'Objectname,PropertyName,NewDataType
'Notes on NewDataType: Can be either Byte Boolean Integer Long Single String
'One more thing to remember all these Properties are case sensetive
    ReDim gComFix(0)
    Open App.Path & "\ComFix.txt" For Input As #1
    Dim data As String
    Dim Temp
    
    Do While Not EOF(1)
        Line Input #1, data
        
        Temp = Split(data, ",")
        gComFix(UBound(gComFix)).ObjectName = Temp(0)
        gComFix(UBound(gComFix)).PropertyName = Temp(1)
        gComFix(UBound(gComFix)).NewType = Temp(2)
        ReDim Preserve gComFix(UBound(gComFix) + 1)
    Loop
    Close #1
    ReDim Preserve gComFix(UBound(gComFix) - 1)

    
End Sub
