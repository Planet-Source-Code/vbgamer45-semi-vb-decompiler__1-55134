Attribute VB_Name = "modPeSkeleton"
Option Explicit

'Common dialog constants
Const cdlOFNReadOnly = &H1
Const cdlOFNOverwritePrompt = &H2
Const cdlOFNHideReadOnly = &H4
Const cdlOFNNoChangeDir = &H8
Const cdlOFNHelpButton = &H10
Const cdlOFNNoValidate = &H100
Const cdlOFNAllowMultiselect = &H200
Const cdlOFNExtensionDifferent = &H400
Const cdlOFNPathMustExist = &H800
Const cdlOFNFileMustExist = &H1000
Const cdlOFNCreatePrompt = &H2000
Const cdlOFNShareAware = &H4000
Const cdlOFNNoReadOnlyReturn = &H8000
Const cdlOFNNoLongNames = &H40000
Const cdlOFNExplorer = &H80000
Const cdlOFNNoDereferenceLinks = &H100000
Const cdlOFNLongNames = &H200000
Const cdlCancel = 32755

'File constants
Public Const MAXNUMBERIMAGEENTRIES = 16  'Current PE file limit for number of DataDirectory directories
Public Const DOS_SIGNATURE = 23117    '"MZ" = 0x4D5A
Public Const PE_SIGNATURE = 17744    '"PE" + 0x00 = 0x50450000
Public Const LENNAME = 8              'Length of Section Header names
Public Const MAXSECTIONS = 16        'Current PE file limit for number of SectionHeader sections
Public Const VBVERTEXT = "VB5!"      'Current compiler/linker output version
Public Const VBINTRPTRTEXT = "MSVBVM60.DLL"  'Current interpreter filename

'Application-wide variables
Global SFile As String         'Source file name
Global SFilePath As String     'Source file path
Global DFile As String         'Destination directory dummy filename
Global DDirPath As String      'Destination directory path
Global ErrorFlag As Boolean    'Generic error flag
Global InFileNumber As Integer
Global OutFileNumber As Integer
Global DecLoadOffset As Double 'App load address

'--------------------------Begin Object structures-----------------
Public Type VB_Signature                    'VB info
    VBVer As String * 4                     'BYTE * 4 = compiler/linker version
    VBIntrptr As String * 12                'BYTE * 12 = interpreter filename
End Type

'General VB6 application data
Public Type App_Data                    'Application specific data for RACE
    DosHeaderOffset As Integer       'DWORD = location of DOS header (offset)in app
    PeHeaderOffset As Double         'DWORD = location of PE header (offset)in app
    OptHeaderOffset As Double        'DWORD = location of OPT header (offset)in app
    SecHeaderOffset As Double        'DWORD = location of SEC header (offset)in app
    VBStartOffset As Double          'DWORD = location of VB application (offset)in app
    VBVerOffsetRaw As Double         'DWORD = location of VB signature (address) in app
    VBVerOffsetMasked As Double      'DWORD = location of VB signature (offset)in app
    VBIntrptrOffset As Double        'DWORD = location of VB interpreter (offset)in app
    ProjRscDataOffset As Double      'DWORD = location of Resource Data (offset)in app
    ProjDataAppReference As Double   'DWORD = location of RACE reference (offset) in app
    ProjDataModuleTable As Double    'DWORD = location of VB modules
    ProjVerDataOffsetRaw As Double   'DWORD = address of Version Data Pool in exe
    ProjVerDataOffsetMasked As Double 'DWORD = offset to Version Data Pool in app
    AppModuleCount As Byte           'BYTE = number of VB modules in app
    StartUpName As String            'Name of first referenced object (if any)
    StartUpType As Double            'Signature of startup object
    StartUpOffset As Double          'Address of object block start
    FormTableAddress As Double       'Address of form list table
    BasicTableAddress As Double      'Address of basic list table
    CompileType As String            'PCODE or NCODE
    VBProjectHeaderOffset As Double  'DWORD = offset in exe of Project signature
    NativeCodeAddressRaw As Double   'DWORD 0x0 if PCODE, address if NCODE
    NativeCodeAddressMasked As Double 'DWORD 0x0 if PCODE, offset in exe if NCODE
End Type

'General VB6 application data
Public Type VBStart_Header
    PushStartOpcode As Integer      'BYTE = opcode 68 = push
    PushStartAddress As Double      'DWORD = address of VB signature
    CallStartOpcode As Integer      'BYTE = opcode E8 = jmp
    CallStartAddress As Double      'DWORD = address of interpreter entry
End Type

'Generic DOS file data
Public Type Dos_Header              'Standard DOS header
    Magic As Double                 'WORD
    NumBytesLastPage As Double      'WORD
    NumPages As Double              'WORD
    NumRelocates As Double          'WORD
    NumHeaderBlks As Double         'WORD
    NumMinBlks As Double            'WORD
    NumMaxBlks As Double            'WORD
    SSPointer As Double             'WORD
    SPPointer As Double             'WORD
    Checksum As Double              'WORD
    IPPointer As Double             'WORD
    CurrentSeg As Double            'WORD
    RelocTablePointer As Double     'WORD
    Overlay As Double               'WORD
    ReservedW1 As Double            'WORD
    ReservedW2 As Double            'WORD
    ReservedW3 As Double            'WORD
    ReservedW4 As Double            'WORD
    OEMType As Double               'WORD
    OEMData As Double               'WORD
    ReservedW5 As Double            'WORD
    ReservedW6 As Double            'WORD
    ReservedW7 As Double            'WORD
    ReservedW8 As Double            'WORD
    ReservedW9 As Double            'WORD
    ReservedW10 As Double           'WORD
    ReservedW11 As Double           'WORD
    ReservedW12 As Double           'WORD
    ReservedW13 As Double           'WORD
    ReservedW14 As Double           'WORD
    ExeHeaderPointer As Double      'DWORD
End Type

'PE file data
Public Type PE_Header               'Standard PE header
    Magic As Double                 'DWORD
    Machine As Double               'WORD
    NumSections As Double           'WORD
    TimeDate As Double              'DWORD
    SymbolTablePointer As Double    'DWORD
    NumSymbols As Double            'DWORD
    OptionalHdrSize As Double       'WORD
    Properties As Double            'WORD
End Type

'PE file data
Public Type Data_Dir                'Standard Data Directory
    Name As String                  'Variable
    Address As Double               'DWORD
    Size As Double                  'DWORD
End Type

'PE file data
Public Type Opt_Header              'Standard Option Header
    Magic As Double                 'WORD
    MajLinkerVer As Integer         'BYTE
    MinLinkerVer As Integer         'BYTE
    CodeSize As Double              'DWORD
    InitDataSize As Double          'DWORD
    UninitDataSize As Double        'DWORD
    EntryPoint As Double            'DWORD
    CodeBase As Double              'DWORD
    DataBase As Double              'DWORD
    ImageBase As Double             'DWORD
    SectionAlignment As Double      'DWORD
    FileAlignment As Double         'DWORD
    MajOSVer As Double              'WORD
    MinOSVer As Double              'WORD
    MajImageVer As Double           'WORD
    MinImageVer As Double           'WORD
    MajSSysVer As Double            'WORD
    MinSSysVer As Double            'WORD
    Win32Ver As Double              'DWORD
    SizeImage As Double             'DWORD
    SizeHeader As Double            'DWORD
    Checksum As Double              'DWORD
    SSystem As Double               'WORD
    DLLProperties As Double         'WORD
    SSizeRes As Double              'DWORD
    SSizeCom As Double              'DWORD
    HSizeRes As Double              'DWORD
    HSizeCom As Double              'DWORD
    LFlags As Double                'DWORD
    NumRVA_Sizes As Double          'DWORD
    DataDirectory(MAXNUMBERIMAGEENTRIES) As Data_Dir
End Type

'PE file data
Public Type Sec_Header              'Standard Section Header
    SecName As String * LENNAME     'BYTE[LENNAME]
    Misc As Double                  'DWORD
    Address As Double               'DWORD
    SizeRawData As Double           'DWORD
    RawDataPointer As Double        'DWORD
    RelocationPointer As Double     'DWORD
    LineNumPointer As Double        'DWORD
    NumRelocations As Double        'WORD
    NumLineNumbers As Double        'WORD
    Properties As Double            'DWORD
End Type

Public AppData As App_Data
Public VBStartHeader As VBStart_Header
Public VBSignature As VB_Signature
Public DosHeader As Dos_Header
Public PeHeader As PE_Header
Public OptHeader As Opt_Header
Public SecHeader(MAXSECTIONS) As Sec_Header




Public Function CheckHeader() As Boolean

    'Assume a good file
    CheckHeader = True
        
    '************************
    'All files must start with the DOS signature
    '************************
       
    'Save the DOSHeader offset (always 0x0000!)
    AppData.DosHeaderOffset = 0
    
    'Get the DOS signature, check for error
    Call GetDOSSignature
    
    If ErrorFlag = True Then
        CheckHeader = False
        Exit Function
    End If
        
    '*******************************
    'The DOS header follows the DOS signature
    '*******************************
     
    'Get the Dos header, check for error
    Call GetDOSHeader
    
    If ErrorFlag = True Then
        CheckHeader = False
        Exit Function
    End If
       
    '*******************************
    'The DOS header holds the PE file signature
    '*******************************
         
    'Move to the location where the PE signature should be
    Seek #InFileNumber, DosHeader.ExeHeaderPointer + 1
    
    'Get the PE signature, check error
    Call GetPESignature
                      
    If ErrorFlag = True Then
        CheckHeader = False
        Exit Function
    End If
    
    '************************************
    'The PEfile header data exists just after the PE signature
    '************************************
   
    'Save the PEHeader offset
    AppData.PeHeaderOffset = Seek(InFileNumber) - 1
    
    'Get the PEHeader
    Call GetPEHeader
      
    '************************************
    'The OPTion header exists just after the PE header
    '***********************************
      
    'Save the OPTHeader offset
    AppData.OptHeaderOffset = Seek(InFileNumber) - 1
        
    'Get the OPTHeader
    Call GetPEOptionHeader
      
    '***********************************
    'The SECtion headers exist just after the option header
    '***********************************
       
    'Save the SECtionHeader offset
    AppData.SecHeaderOffset = Seek(InFileNumber) - 1
     
    'Get the SecHeader
    Call GetPESecHeader
    
    'These sections are not included; they're
    'not needed for VB6 analysis, but could be
    'added if more PE file analysis is desired:
        'DebugDirectory
        'ResourceSection
        'ImportsSection
        'ExportsSection

    '****************************
    'Start the VB app analysis
    '****************************
        
    'Calculate the load offset mask
    DecLoadOffset# = OptHeader.ImageBase
            
    '**************************************
    'The VB Startheader holds the jump vector
    '**************************************
       
    'Get the APP data VB app start location = OPTHeader.EntryPoint
    AppData.VBStartOffset = OptHeader.EntryPoint
        
    'Point file at the VB code start position
    Seek #InFileNumber, AppData.VBStartOffset + 1
    
    'Get the VBStartHeader, check error
    Call GetVBStartHeader
    
    If ErrorFlag = True Then
        CheckHeader = False
        Exit Function
    End If
            
    '**************************************
    'The VB start vector holds the compiler signature
    '**************************************
    
    'Get the APP data VB signature offset
    AppData.VBVerOffsetRaw = VBStartHeader.PushStartAddress
    
    'Calculate the APP offset
    AppData.VBVerOffsetMasked = AppData.VBVerOffsetRaw - DecLoadOffset#
    
    'Point file at the VB signature position
    Seek #InFileNumber, AppData.VBVerOffsetMasked + 1
    
    'Check for VB version (compiler) of this file, check error
    Call GetVBVer
        
    If ErrorFlag = True Then
        CheckHeader = False
        Exit Function
    End If
    
    'Assign this location to our reference
    AppData.ProjDataAppReference = AppData.VBVerOffsetMasked
           
    '*****************************
    'Check if the interpreter name exists
    '*****************************
    
    'Point file at the Data Directory #1 position
    Seek #InFileNumber, OptHeader.DataDirectory(1).Address + 1
    
    'Move ahead 12 bytes
    Seek #InFileNumber, Seek(InFileNumber) + 12
    
    'Get the APP data interpreter address offset
    AppData.VBIntrptrOffset = GetDWord()
    
    'Move to the interpreter signature
    Seek #InFileNumber, AppData.VBIntrptrOffset + 1
    
    'Get the interpreter
    Call GetVBIntrptr
    
    If ErrorFlag = True Then
        CheckHeader = False
        Exit Function
    End If
    
    'If we got here, this is definitely a valid VB6 app
    
End Function

Public Sub GetDOSSignature()

    'Get the first two characters
    DosHeader.Magic = GetWord()
    
    'Check for error
    If DosHeader.Magic <> DOS_SIGNATURE Then
        ErrorFlag = True
    End If
    
End Sub

Public Sub GetDOSHeader()

    'Get DOS header data
        DosHeader.NumBytesLastPage = GetWord()
        DosHeader.NumPages = GetWord()
        DosHeader.NumRelocates = GetWord()
        DosHeader.NumHeaderBlks = GetWord()
        DosHeader.NumMinBlks = GetWord()
        DosHeader.NumMaxBlks = GetWord()
        DosHeader.SSPointer = GetWord()
        DosHeader.SPPointer = GetWord()
        DosHeader.Checksum = GetWord()
        DosHeader.IPPointer = GetWord()
        DosHeader.CurrentSeg = GetWord()
        DosHeader.RelocTablePointer = GetWord()
        DosHeader.Overlay = GetWord()
        DosHeader.ReservedW1 = GetWord()
        DosHeader.ReservedW2 = GetWord()
        DosHeader.ReservedW3 = GetWord()
        DosHeader.ReservedW4 = GetWord()
        DosHeader.OEMType = GetWord()
        DosHeader.OEMData = GetWord()
        DosHeader.ReservedW5 = GetWord()
        DosHeader.ReservedW6 = GetWord()
        DosHeader.ReservedW7 = GetWord()
        DosHeader.ReservedW8 = GetWord()
        DosHeader.ReservedW9 = GetWord()
        DosHeader.ReservedW10 = GetWord()
        DosHeader.ReservedW11 = GetWord()
        DosHeader.ReservedW12 = GetWord()
        DosHeader.ReservedW13 = GetWord()
        DosHeader.ReservedW14 = GetWord()
        DosHeader.ExeHeaderPointer = GetDWord()
        
        'Make sure the potential PE signature location seems reasonable
        If ((DosHeader.ExeHeaderPointer > 4096) Or (DosHeader.ExeHeaderPointer < 64)) Then
            ErrorFlag = True
        End If
        
End Sub

Public Sub GetPESignature()

    'Get the first two characters
    PeHeader.Magic = GetDWord()
        
    'Check for error
    If PeHeader.Magic <> PE_SIGNATURE Then
        ErrorFlag = True
    End If
    
End Sub

Public Sub GetPEOptionHeader()

    'Now get the "optional" header data
    OptHeader.Magic = GetWord()
    OptHeader.MajLinkerVer = GetByte()
    OptHeader.MinLinkerVer = GetByte()
    OptHeader.CodeSize = GetDWord()
    OptHeader.InitDataSize = GetDWord()
    OptHeader.UninitDataSize = GetDWord()
    OptHeader.EntryPoint = GetDWord()
    OptHeader.CodeBase = GetDWord()
    OptHeader.DataBase = GetDWord()
    OptHeader.ImageBase = GetDWord()
    OptHeader.SectionAlignment = GetDWord()
    OptHeader.FileAlignment = GetDWord()
    OptHeader.MajOSVer = GetWord()
    OptHeader.MinOSVer = GetWord()
    OptHeader.MajImageVer = GetWord()
    OptHeader.MinImageVer = GetWord()
    OptHeader.MajSSysVer = GetWord()
    OptHeader.MinSSysVer = GetWord()
    OptHeader.Win32Ver = GetDWord()
    OptHeader.SizeImage = GetDWord()
    OptHeader.SizeHeader = GetDWord()
    OptHeader.Checksum = GetDWord()
    OptHeader.SSystem = GetWord()
    OptHeader.DLLProperties = GetWord()
    OptHeader.SSizeRes = GetDWord()
    OptHeader.SSizeCom = GetDWord()
    OptHeader.HSizeRes = GetDWord()
    OptHeader.HSizeCom = GetDWord()
    OptHeader.LFlags = GetDWord()
    OptHeader.NumRVA_Sizes = GetDWord()
    OptHeader.DataDirectory(0).Name = "EXPORT"
    OptHeader.DataDirectory(0).Address = GetDWord()
    OptHeader.DataDirectory(0).Size = GetDWord()
    OptHeader.DataDirectory(1).Name = "IMPORT"
    OptHeader.DataDirectory(1).Address = GetDWord()
    OptHeader.DataDirectory(1).Size = GetDWord()
    OptHeader.DataDirectory(2).Name = "RESOURCE"
    OptHeader.DataDirectory(2).Address = GetDWord()
    OptHeader.DataDirectory(2).Size = GetDWord()
    OptHeader.DataDirectory(3).Name = "EXCEPTION"
    OptHeader.DataDirectory(3).Address = GetDWord()
    OptHeader.DataDirectory(3).Size = GetDWord()
    OptHeader.DataDirectory(4).Name = "SECURITY"
    OptHeader.DataDirectory(4).Address = GetDWord()
    OptHeader.DataDirectory(4).Size = GetDWord()
    OptHeader.DataDirectory(5).Name = "BASERELOC"
    OptHeader.DataDirectory(5).Address = GetDWord()
    OptHeader.DataDirectory(5).Size = GetDWord()
    OptHeader.DataDirectory(6).Name = "DEBUG"
    OptHeader.DataDirectory(6).Address = GetDWord()
    OptHeader.DataDirectory(6).Size = GetDWord()
    OptHeader.DataDirectory(7).Name = "COPYRIGHT"
    OptHeader.DataDirectory(7).Address = GetDWord()
    OptHeader.DataDirectory(7).Size = GetDWord()
    OptHeader.DataDirectory(8).Name = "GLOBALPTR"
    OptHeader.DataDirectory(8).Address = GetDWord()
    OptHeader.DataDirectory(8).Size = GetDWord()
    OptHeader.DataDirectory(9).Name = "TLS"
    OptHeader.DataDirectory(9).Address = GetDWord()
    OptHeader.DataDirectory(9).Size = GetDWord()
    OptHeader.DataDirectory(10).Name = "LOAD_CONFIG"
    OptHeader.DataDirectory(10).Address = GetDWord()
    OptHeader.DataDirectory(10).Size = GetDWord()
    OptHeader.DataDirectory(11).Name = "unused"
    OptHeader.DataDirectory(11).Address = GetDWord()
    OptHeader.DataDirectory(11).Size = GetDWord()
    OptHeader.DataDirectory(12).Name = "unused"
    OptHeader.DataDirectory(12).Address = GetDWord()
    OptHeader.DataDirectory(12).Size = GetDWord()
    OptHeader.DataDirectory(13).Name = "unused"
    OptHeader.DataDirectory(13).Address = GetDWord()
    OptHeader.DataDirectory(13).Size = GetDWord()
    OptHeader.DataDirectory(14).Name = "unused"
    OptHeader.DataDirectory(14).Address = GetDWord()
    OptHeader.DataDirectory(14).Size = GetDWord()
    OptHeader.DataDirectory(15).Name = "unused"
    OptHeader.DataDirectory(15).Address = GetDWord()
    OptHeader.DataDirectory(15).Size = GetDWord()
    
End Sub

Public Sub GetPEHeader()

    'Fill up the PE header structure
    PeHeader.Machine = GetWord()
    PeHeader.NumSections = GetWord()
    PeHeader.TimeDate = GetDWord()
    PeHeader.SymbolTablePointer = GetDWord()
    PeHeader.NumSymbols = GetDWord()
    PeHeader.OptionalHdrSize = GetWord()
    PeHeader.Properties = GetWord()
        
End Sub

Public Sub GetPESecHeader()

    Dim counter As Integer
    Dim Counter2 As Integer
        
    'Name = SecName
    'Address = RVA
    'RawDataPointer = Offset
    'Size = SizeRawData
    'Flags = Properties
    
    'All names are max 8 chars, starting with "."; get all 8 chars, remove the
    'trailing blanks, loop back until all sections done
    For counter = 1 To PeHeader.NumSections
        SecHeader(counter).SecName = Input(LENNAME, #InFileNumber%)
        For Counter2 = 1 To LENNAME
            'Remove the trailing blanks
            If Asc(Mid$(SecHeader(counter).SecName, Counter2, 1)) = 0 Then
                Mid$(SecHeader(counter).SecName, Counter2, 1) = ""
            End If
        Next
        
        'Fill in the rest of the structure
        SecHeader(counter).Misc = GetDWord()
        SecHeader(counter).Address = GetDWord()
        SecHeader(counter).SizeRawData = GetDWord()
        SecHeader(counter).RawDataPointer = GetDWord()
        SecHeader(counter).RelocationPointer = GetDWord()
        SecHeader(counter).LineNumPointer = GetDWord()
        SecHeader(counter).NumRelocations = GetWord()
        SecHeader(counter).NumLineNumbers = GetWord()
        SecHeader(counter).Properties = GetDWord()
    Next
    
End Sub

Public Sub GetVBStartHeader()

    'All VB start headers are a "Push" (5 bytes), followed by
    'a "Call" (5 bytes)
    VBStartHeader.PushStartOpcode = GetByte()
    VBStartHeader.PushStartAddress = GetDWord()
    VBStartHeader.CallStartOpcode = GetByte()
    VBStartHeader.CallStartAddress = GetDWord()
   
    'This should be the hex code "68 xx xx xx xx"
    If VBStartHeader.PushStartOpcode <> ConvertHex("68") Then
        ErrorFlag = True
        Exit Sub
    End If
        
    'This should be the hex code "E8 xx xx xx xx"
    If VBStartHeader.CallStartOpcode <> ConvertHex("E8") Then
        ErrorFlag = True
    End If
    
End Sub

Public Sub GetVBVer()

    'Fill in the VBSignature structure
    Mid$(VBSignature.VBVer$, 1) = Chr$(GetByte())
    Mid$(VBSignature.VBVer$, 2) = Chr$(GetByte())
    Mid$(VBSignature.VBVer$, 3) = Chr$(GetByte())
    Mid$(VBSignature.VBVer$, 4) = Chr$(GetByte())
    
    'The version should be "VB5!"
    If VBSignature.VBVer$ <> VBVERTEXT Then
        ErrorFlag = True
    End If
    
End Sub

Public Sub GetVBIntrptr()

    'Get VB interpreter name
    VBSignature.VBIntrptr$ = GetDosString()
    
    'The interpreter should be "MSVBVM60.DLL" or "MSVBVM50.DLL"
    If VBSignature.VBIntrptr$ <> VBINTRPTRTEXT Then
        ErrorFlag = True
    End If
    
End Sub

Public Function GetDWord() As Double
    
   GetDWord# = GetWord()
   GetDWord# = GetDWord# + 65536# * GetWord()
      
End Function

Public Function GetWord() As Double

    GetWord# = GetByte()
    GetWord# = GetWord# + 256# * GetByte()
       
End Function

Public Function GetByte() As Byte

    Dim DataByte As Byte
    
    'Read the data
    Get #InFileNumber, , DataByte
    
    'Return it
    GetByte = DataByte
      
End Function


Public Function ConvertHex(HexData As String) As Double

    Dim count As Integer
    
    'Get HexData as a 8 byte leading zero string
    HexData$ = "0000000" & HexData$
    HexData$ = Right$(HexData$, 8)
    
    'Sum up powers of 16 for 8 byte values
    ConvertHex = 0#
    For count = 1 To 8
        ConvertHex = ConvertHex + 16 ^ (count - 1) * (HexToDec(Mid$(HexData$, 9 - count, 1)))
    Next
    
End Function


Public Function GetDosString() As String

    'Clear string
    GetDosString$ = ""
     
    'Get a Dos string char by char
    Do
        'Get a char
        GetDosString$ = GetDosString$ & Chr$(GetByte())
        
    'Continue until we get a "0"
    Loop Until (Asc(Right$(GetDosString$, 1)) = 0)
        
    'Remove trailing zero
    GetDosString$ = Left$(GetDosString$, Len(GetDosString$) - 1)
  
End Function


Public Function HexToDec(HexDigit As String) As Double

    'Simple brute force hex-to-decimal conversion
    If HexDigit = "F" Then
        HexToDec# = 15#
    ElseIf HexDigit = "E" Then
        HexToDec# = 14#
    ElseIf HexDigit = "D" Then
        HexToDec# = 13#
    ElseIf HexDigit = "C" Then
        HexToDec# = 12#
    ElseIf HexDigit = "B" Then
        HexToDec# = 11#
    ElseIf HexDigit = "A" Then
        HexToDec# = 10#
    Else
        HexToDec# = Val(HexDigit$)
    End If
    
End Function

