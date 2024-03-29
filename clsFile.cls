VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
Private mFileNumber As Integer ' fileNumber
Private mShortFileName As String
Private mFileName As String
Private mOffset As Long



Public Sub Setup(sFileName As String)
On Local Error GoTo localError
    If Len(mFileName) Then
        Debug.Print "clsFile.FileName", "Can not change filename once it has been set."
        Exit Sub
    End If
    
    ' check if its a valid file
    Debug.Print "Initializing File"
    Dim sExtention As String
    sExtention = LCase(Right(sFileName, 4))
    If Len(Dir(sFileName)) = 0 Then
        Debug.Print "File Not Found.", "clsFile.Initialize"
        Exit Sub
'    ElseIf (sExtention <> ".exe" And sExtention <> ".dll" And sExtention <> ".ocx") Then
'        Root.Display.Error "Extention not supported.", "clsFile.Initialize"
'        Exit Sub
    End If
    ' Find a free file number
    mFileNumber = FreeFile
    'set pathnames
    mFileName = sFileName
    mShortFileName = Dir(mFileName) 'trim path name
    ' Open the file for reading
    Open mFileName For Binary Access Read Lock Write As #mFileNumber
    Exit Sub
localError:
    Dim errDesc As String
    errDesc = Err.Description
    Err.Clear
    Debug.Print errDesc, "clsFile.Initialize()"

End Sub
Public Function GetGUID(Optional Offset As Long = -1) As String
On Local Error GoTo localError
Dim sTemp As String
If Offset = -1 Then Offset = mOffset
mOffset = mOffset + 16
sTemp = sHexStringFromString(GetString(Offset, 16), False)
GetGUID = "{" & Mid(sTemp, 7, 2) & Mid(sTemp, 5, 2) & Mid(sTemp, 3, 2) & Mid(sTemp, 1, 2) & "-" & Mid(sTemp, 11, 2) & Mid(sTemp, 9, 2) & "-" & Mid(sTemp, 15, 2) & Mid(sTemp, 13, 2) & "-" & Mid(sTemp, 17, 4) & "-" & Mid(sTemp, 21, 12) & "}"
        Exit Function
localError:
    Dim errDesc As String
    errDesc = Err.Description
    Err.Clear
    Debug.Print errDesc, "clsFile.GetGUID()"
End Function
Public Function GetByte(Optional Offset As Long = -1) As Byte
On Local Error GoTo localError
If Offset = -1 Then Offset = mOffset
mOffset = mOffset + 1
Get mFileNumber, Offset + 1, GetByte
        Exit Function
localError:
    Dim errDesc As String
    errDesc = Err.Description
    Err.Clear
    Debug.Print errDesc, "clsFile.GetByte()"
End Function
Public Function GetBytes(Length As Integer, Optional Offset As Long = -1) As Byte()
On Local Error GoTo localError
ReDim GetBytes(Length - 1)
If Offset = -1 Then Offset = mOffset
mOffset = mOffset + Length - 1
    Get mFileNumber, Offset + 1, GetBytes
        Exit Function
localError:
    Dim errDesc As String
    errDesc = Err.Description
    Err.Clear
    Debug.Print errDesc, "clsFile.GetBytes()"
End Function
Public Function GetString(Optional Offset As Long = -1, Optional ByVal Length As Long = -1, Optional Unicode As Boolean = False) As String
On Local Error GoTo localError
If Offset = -1 Then Offset = mOffset
    Dim b As Byte
    Dim i As Integer
    Seek mFileNumber, Offset + 1     'goto new loc
    Do
        If Unicode = True Then
            Get mFileNumber, , i
            b = (i And &HFF) 'convert unicode to regular
        Else
            Get mFileNumber, , b
        End If
        GetString = GetString & Chr$(b)
    Loop Until (Length = -1 And b = 0) Or (Len(GetString) = Length)
    If Len(GetString) <> Length Then GetString = Left(GetString, Len(GetString) - 1)  'trim null
    If Length = -1 Then
        mOffset = mOffset + Len(GetString) + 1
    Else
        mOffset = mOffset + Length
    End If
    Exit Function
localError:
    Dim errDesc As String
    errDesc = Err.Description
    Err.Clear
    Debug.Print errDesc, "clsFile.GetString()"
End Function
Public Function GetInteger(Optional Offset As Long = -1) As Integer
On Local Error GoTo localError
If Offset = -1 Then Offset = mOffset
mOffset = mOffset + 2

    Get mFileNumber, Offset + 1, GetInteger
        Exit Function
localError:
    Dim errDesc As String
    errDesc = Err.Description
    Err.Clear
    Debug.Print errDesc, "clsFile.GetInteger()"
End Function
Public Function GetLong(Optional Offset As Long = -1) As Long
On Local Error GoTo localError
'Dim l As Long
If Offset = -1 Then Offset = mOffset
mOffset = mOffset + 4
    
    Get mFileNumber, Offset + 1, GetLong
        Exit Function
localError:
    Dim errDesc As String
    errDesc = Err.Description
    Err.Clear
    Debug.Print errDesc, "clsFile.GetLong()"
End Function
Public Function GetDouble(Optional Offset As Long = -1) As Double
On Local Error GoTo localError
'Dim l As Long
If Offset = -1 Then Offset = mOffset
mOffset = mOffset + 4
    
    Get mFileNumber, Offset + 1, GetDouble
        Exit Function
localError:
    Dim errDesc As String
    errDesc = Err.Description
    Err.Clear
    Debug.Print errDesc, "clsFile.GetDouble()"
End Function
Public Function GetSingle(Optional Offset As Long = -1) As Single
On Local Error GoTo localError

    If Offset = -1 Then Offset = mOffset
    mOffset = mOffset + 4
    
    Get mFileNumber, Offset + 1, GetSingle
    Exit Function
localError:
    Dim errDesc As String
    errDesc = Err.Description
    Err.Clear
    Debug.Print errDesc, "clsFile.GetSingle()"
End Function
Public Function GetCurrency(Optional Offset As Long = -1) As Currency
On Local Error GoTo localError
If Offset = -1 Then Offset = mOffset
mOffset = mOffset + 1
Get mFileNumber, Offset + 1, GetCurrency
        Exit Function
localError:
    Dim errDesc As String
    errDesc = Err.Description
    Err.Clear
    Debug.Print errDesc, "clsFile.GetCurrency()"
End Function

Private Sub Class_Terminate()
   Close #mFileNumber
End Sub
Property Get Length() As Long:    Length = LOF(mFileNumber): End Property
Property Get FileNumber() As Integer:    FileNumber = mFileNumber: End Property
Property Get FileName() As String:    FileName = mFileName: End Property
Property Get ShortFileName() As String:    ShortFileName = mShortFileName: End Property
