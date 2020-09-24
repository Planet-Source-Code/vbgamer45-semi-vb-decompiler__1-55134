Attribute VB_Name = "modControls"
'Used if COM doesn't work
'Control Sepeartor Constatns
Public Const vbFormNewChildControl = 511 'FF01
Public Const vbFormExistingChildControl = 767 'FF02
Public Const vbFormChildControl = 1023 'FF03
Public Const vbFormEnd = 1279 'FF04
Public Const vbFormMenu = 1535 'FF05
'Control Header
Public Type ControlHeader
    Length As Integer
    unknown As Integer
    un1 As Byte
    cName As String
    un2 As Byte
    cType As Byte
End Type
Public Type ControlSize
    clientLeft As Integer
    un1 As Integer
    clientTop As Integer
    un2 As Integer
    clientWidth As Integer
    un3 As Integer
    clientHeight As Integer
    un4 As Integer
End Type
'Used in cType
Public Enum ControlType
    vbPictureBox = 0
    vbLabel = 1
    vbTextbox = 2
    vbFrame = 3
    vbCommandbutton = 4
    vbCheckbox = 5
    vbOptionbutton = 6
    vbCombobox = 7
    vbListbox = 8
    vbHscroll = 9
    vbVscroll = 10
    vbTimer = 11
    vbForm = 13
    vbDriveListbox = 16
    vbDirectoryListbox = 17
    vbFileListbox = 18
    vbMenu = 19
    vbMDIForm = 20
    vbShape = 22
    vbLine = 23
    vbImage = 24
    vbData = 37
    vbOLE = 38
    vbUserControl = 40
    vbPropertyPage = 41
    vbUserDocument = 42
End Enum






