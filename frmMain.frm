VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   Caption         =   "Semi VB Decompiler by vbgamer45"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7905
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1                   vbgamer45"
   ScaleHeight     =   6495
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Tag             =   "                                   v b g a m e r 4 5"
   Begin VB.TextBox txtFinal 
      Height          =   1695
      Index           =   0
      Left            =   7920
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.ListBox lstTypeInfos 
      Height          =   2400
      Left            =   7920
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.ListBox lstMembers 
      Height          =   2400
      Left            =   7920
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   3720
      Visible         =   0   'False
      Width           =   3975
   End
   Begin RichTextLib.RichTextBox txtFunctions 
      Height          =   615
      Left            =   8400
      TabIndex        =   6
      Top             =   5640
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1085
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":27A2
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   5
      Top             =   6225
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   476
      Style           =   1
      SimpleText      =   "Credits: VB Decompiling community... Sarge, Mr. Unleaded, Moogman, Warning, and others..."
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog Cd1 
      Left            =   5400
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imglistControl 
      Left            =   60
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   43
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2824
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2DBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3358
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":43AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4744
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5796
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5AE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5E3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":618C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":64DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6830
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6B84
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6ED8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":722A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":757C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":78CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7C20
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7F72
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":82C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8616
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8968
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8CBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":900E
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9360
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":96B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9A04
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9D56
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A0A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A91A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B18C
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B9FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C270
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CAE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D354
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D6A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DF18
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E78A
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EFFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F350
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F6A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FC3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":101D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10770
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvProject 
      Height          =   6075
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   10716
      _Version        =   393217
      Indentation     =   617
      LabelEdit       =   1
      Style           =   7
      HotTracking     =   -1  'True
      ImageList       =   "imglistControl"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin TabDlg.SSTab sstViewFile 
      Height          =   6075
      Left            =   3480
      TabIndex        =   1
      Tag             =   "T{20/21/}"
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   10716
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Code"
      TabPicture(0)   =   "frmMain.frx":10D0A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtCode"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Properties"
      TabPicture(1)   =   "frmMain.frx":10D26
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fxgEXEInfo"
      Tab(1).ControlCount=   1
      Begin RichTextLib.RichTextBox txtCode 
         Height          =   5535
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   9763
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmMain.frx":10D42
      End
      Begin MSFlexGridLib.MSFlexGrid fxgEXEInfo 
         Height          =   5535
         Left            =   -74940
         TabIndex        =   2
         Top             =   360
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   9763
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   -2147483627
         ForeColorFixed  =   12829635
         GridColorFixed  =   8421504
         ScrollTrack     =   -1  'True
         GridLinesFixed  =   1
         AllowUserResizing=   1
      End
      Begin MSComDlg.CommonDialog cdlShow 
         Left            =   -74940
         Top             =   7590
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList imlIcons 
         Left            =   -74940
         Top             =   7590
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":10DC4
               Key             =   "COCLASS"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":11216
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":11668
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":11ABA
               Key             =   "i4"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":11F0C
               Key             =   "i2"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":12066
               Key             =   "i0"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":121C0
               Key             =   "i1"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1231A
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":12474
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtbFile 
         Height          =   8535
         Left            =   -74940
         TabIndex        =   3
         Top             =   360
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   15055
         _Version        =   393217
         Enabled         =   -1  'True
         HideSelection   =   0   'False
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         OLEDragMode     =   0
         OLEDropMode     =   1
         TextRTF         =   $"frmMain.frx":12A0E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtResult 
         Height          =   675
         Left            =   -74940
         TabIndex        =   4
         Top             =   8190
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   1191
         _Version        =   393217
         BackColor       =   12632256
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmMain.frx":12A98
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Members:"
      Height          =   195
      Left            =   7920
      TabIndex        =   10
      Top             =   3480
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TypeInfos:"
      Height          =   195
      Left            =   7920
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileGenerate 
         Caption         =   "&Generate vbp"
         Enabled         =   0   'False
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuFileExportMemoryMap 
         Caption         =   "&Export Memory Map"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileRecent1 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileRecent2 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileRecent3 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileRecent4 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'###########################################
'#VB Semi Decompiler
'#By vbgamer45
'#Credits:
'#Some code from decompiler.theautomaters.com  The VB Decompiling Community
'#Sarge for the PE Skeleton
'#Mr. Unleaded for MemoryMap
'#Moogman for TypeViewer
'#And from Warning for treeview
'###########################################

'The following is used for the browse for folder dialog
Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type
Const BIF_RETURNONLYFSDIRS = 1
Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Sub Form_Load()
    Me.Caption = " Semi VB Decompiler by vbgamer45 Version: " & Version
    Call PrintReadMe
    'Setup Variables
    gSkipCom = False
    gDumpData = False
    
    'Get the recent file list
    Dim Recent1Title As String
    Dim Recent2Title As String
    Dim Recent3Title As String
    Dim Recent4Title As String
    
    Recent1Title = GetSetting("VB Decompiler", "Options", "Recent1FileTitle", "")
    Recent2Title = GetSetting("VB Decompiler", "Options", "Recent2FileTitle", "")
    Recent3Title = GetSetting("VB Decompiler", "Options", "Recent3FileTitle", "")
    Recent4Title = GetSetting("VB Decompiler", "Options", "Recent4FileTitle", "")
    If Recent1Title <> "" Then
        mnuFileRecent1.Visible = True
        mnuFileSep1.Visible = True
        mnuFileRecent1.Caption = Recent1Title
    End If
    If Recent2Title <> "" Then
        mnuFileRecent2.Visible = True
        mnuFileRecent2.Caption = Recent2Title
    End If
    If Recent3Title <> "" Then
        mnuFileRecent3.Visible = True
        mnuFileRecent3.Caption = Recent3Title
    End If
    If Recent4Title <> "" Then
        mnuFileRecent4.Visible = True
        mnuFileRecent4.Caption = Recent4Title
    End If
    
    'Setup the COM Functions
    Set tliTypeLibInfo = New TypeLibInfo
    'GUID for vb6.olb used to find the gui opcodes of the standard controls
    tliTypeLibInfo.LoadRegTypeLib "{FCFB3D2E-A0FA-1068-A738-08002B3371B5}", 6, 0, 9
    tliTypeLibInfo.LoadRegTypeLib "{FCFB3D2E-A0FA-1068-A738-08002B3371B5}", 6, 0, 9
    Call ProcessTypeLibrary
    tliTypeLibInfo.AppObjString = "<Global>"
    'Load the functions
  '  Call getFunctionsFromFile("C:\Program Files\Microsoft Visual Studio\VB98\VB6.OLB")
    'Load Com Hacks
    Call modGlobals.LoadCOMFIX
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    tvProject.Height = Me.Height - StatusBar1.Height - 700
    sstViewFile.Height = Me.Height - StatusBar1.Height - 700
    
End Sub

Private Sub lstMembers_Click()
 Dim tliInvokeKinds As InvokeKinds
    tliInvokeKinds = lstMembers.ItemData(lstMembers.ListIndex)
 
   ' MsgBox ReturnGuiOpcode(lstTypeInfos.ItemData(lstTypeInfos.ListIndex), tliInvokeKinds, lstMembers.[_Default])
    If lstTypeInfos.ListIndex <> -1 Then
    MsgBox ReturnDataType(lstTypeInfos.ItemData(lstTypeInfos.ListIndex), tliInvokeKinds, lstMembers.[_Default])
    End If
End Sub

Private Sub lstTypeInfos_Click()
    Dim tliTypeInfo As TypeInfo
    Set tliTypeInfo = tliTypeLibInfo.GetTypeInfo(Replace(Replace(lstTypeInfos.List(lstTypeInfos.ListIndex), "<", ""), ">", ""))
    'Use the ItemData in lstTypeInfos to set the SearchData for lstMembers
    tliTypeLibInfo.GetMembersDirect lstTypeInfos.ItemData(lstTypeInfos.ListIndex), lstMembers.hWnd, , , True
    
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileExportMemoryMap_Click()

    Set gVBFile = New clsFile
    
    Call gVBFile.Setup(SFilePath)
    
    Set gMemoryMap = New clsMemoryMap
    hascollision = gMemoryMap.AddSector(VBStartHeader.PushStartAddress - OptHeader.ImageBase, 102, "vb header")
    hascollision = gMemoryMap.AddSector(gVBHeader.aProjectInfo - OptHeader.ImageBase, 572, "project info")
    hascollision = gMemoryMap.AddSector(gProjectInfo.aObjectTable - OptHeader.ImageBase, 84, "objecttable")
    Dim i As Integer
    For i = 0 To gObjectTable.ObjectCount1
    
    Next
    
    gMemoryMap.ExportToHTML 'exports to File.Name & ".html"
    MsgBox "Memory Map Created!"

End Sub

Private Sub mnuFileGenerate_Click()
    Dim sPath As String
    Dim structFolder As BROWSEINFO
    Dim iNull As Integer
    structFolder.hOwner = Me.hWnd
    structFolder.lpszTitle = "Browse for folder"
    structFolder.ulFlags = BIF_RETURNONLYFSDIRS
    Dim Ret As Long
    Ret = SHBrowseForFolder(structFolder)
    If Ret Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList Ret, sPath
        'free the block of memory
        CoTaskMemFree Ret
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    'Write The Project File
    Call WriteVBP(sPath & "\" & ProjectName & ".vbp")
    'Write the forms
    Call WriteForms(sPath & "\")
    'Write the modules
    For i = 0 To UBound(gObject)
        If gObject(i).ObjectType = 98305 Then
           Call modOutput.WriteModules(sPath & "\" & gObjectNameArray(i) & ".bas", gObjectNameArray(i))
        End If
    Next
    'Write the classes
    For i = 0 To UBound(gObject)
        If gObject(i).ObjectType = 1146883 Then
            Call modOutput.WriteClasses(sPath & "\" & gObjectNameArray(i) & ".cls", gObjectNameArray(i))
        End If
    Next
    'Write the user controls
    
    MsgBox "Done"
End Sub

Private Sub mnuFileOpen_Click()
    Cd1.FileName = ""
    Cd1.DialogTitle = "Select VB5/VB6 exe"
    Cd1.Filter = "Executable (*.exe)|*.exe"
    Cd1.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
    Cd1.ShowOpen
    
    If Cd1.FileName = "" Then Exit Sub
    
    Call OpenVBExe(Cd1.FileName, Cd1.FileTitle)
End Sub

Sub OpenVBExe(FilePath As String, FileTitle As String)
Dim bFormEndUsed As Boolean
   
    'Erase existing data
    bFormEndUsed = False
    For textCount = 0 To txtFinal.UBound
        txtFinal(textCount).Text = ""
    Next
    
    SFilePath = ""
    SFile = ""
    ReDim gControlNameArray(0) 'Treeveiw control list
    'clear the nodes
    tvProject.Nodes.Clear
    'Save name and path
    SFilePath = FilePath
    SFile = FileTitle
    
    'Reset the error flag
    ErrorFlag = False
    
    'Get a file handle
    InFileNumber = FreeFile
    
    'Check for error
    'On Error GoTo AnalyzeError
    
    'Access the file
    Open SFilePath For Binary As #InFileNumber
       
    'Is it a VB6 file?
    If CheckHeader() = True Then
        'Good file
        
        Close #InFileNumber
    Else
       'Bad file
       
       MsgBox "Not a VB6 file.", vbOKOnly Or vbCritical Or vbApplicationModal, "Bad file!"
        Close #InFileNumber
        Exit Sub
    End If

    Startoffset = VBStartHeader.PushStartAddress - OptHeader.ImageBase
    
    f = FreeFile
    Open FilePath For Binary Access Read Lock Read As f
        'Goto begining of vb header
        Seek f, Startoffset + 1
        'Get the vb header
        Get #f, , gVBHeader
        
        AppData.FormTableAddress = gVBHeader.aGUITable
        'GetHelpFile
        Seek #f, Startoffset + 1 + gVBHeader.oHelpFile 'Loc(f) + gVBHeader.oHelpFile + 1
        HelpFile = GetUntilNull(f)
        'Get Project Name
        Seek #f, Startoffset + 1 + gVBHeader.oProjectName
        ProjectName = GetUntilNull(f)
        'Project Title
        Seek #f, Startoffset + 1 + gVBHeader.oProjectTitle
        ProjectTitle = GetUntilNull(f)
        'ExeName
        Seek #f, Startoffset + 1 + gVBHeader.oProjectExename
        ProjectExename = GetUntilNull(f)
        'Get Project Description
        Seek #f, gVBHeader.aProjectDescription + 65 - OptHeader.ImageBase
        ProjectDescription = GetUntilNull(f)
        'Get Project Info Table
        Seek f, gVBHeader.aProjectInfo + 1 - OptHeader.ImageBase
        Get #f, , gProjectInfo
        
        'Get External Table
         Seek f, gProjectInfo.aExternalTable + 1 - OptHeader.ImageBase
       
        Get #f, , gExternalTable
        'MsgBox gVBHeader.aExternalComponentTable + 1 - OptHeader.ImageBase
        'Get External Library
       ' redim gexternallibary
        If gProjectInfo.ExternalCount > 0 Then
            Seek f, gExternalTable.aExternalLibrary + 1 - OptHeader.ImageBase
            Get #f, , gExternalLibrary
        End If
  
  
'#################################################
'Process Forms And Control Properties
'################################################
        Seek f, gVBHeader.aGUITable + 1 - OptHeader.ImageBase
        'Get Form table
        If gVBHeader.FormCount > 0 Then
            ReDim gGuiTable(gVBHeader.FormCount - 1)
            Get #f, , gGuiTable
        End If
        Dim fPos As Long 'Holds current location in the file used for controlheader
        Dim cListIndex As Integer ' Used for COM
        Dim cControlHeader As ControlHeader
        
        Dim lForm As Integer
        'Loop though each form...
        For lForm = 0 To UBound(gGuiTable)

        Seek f, gGuiTable(lForm).aFormPointer + 94 - OptHeader.ImageBase
        
'Loop from new child control
NewControl:
        fPos = Loc(f)
        Get #f, , cControlHeader
       If gDumpData = True Then
        Dim fHeaderEnd As Long
        fHeaderEnd = Loc(f)
        Dim fControlEnd As Long
        'Store each object's information in a file
        fControlEnd = DumpObject(f, cControlHeader.cName, cControlHeader.Length, fPos, fHeaderEnd)
        If gSkipCom = True Then
        'Get all dumps of the controls even though COM is off
            If fControlEnd <> -1 Then
                Seek f, fControlEnd
                GoTo NewControl
            End If
        End If
       End If
       If gSkipCom = False Then
        Dim tliTypeInfo As TypeInfo 'Used for COM to find information about the properties of the control
        Dim FileLen As Long 'Used to caculate how much father to go in the control
        'Select what type of control it is
        Select Case cControlHeader.cType
            Case vbPictureBox '= 0
                cListIndex = 22
                Call AddText("Begin VB.PictureBox " & cControlHeader.cName)
            Case vbLabel '= 1
                cListIndex = 14
                Call AddText("Begin VB.Label " & cControlHeader.cName)
            Case vbTextbox ' = 2
                cListIndex = 27
                Call AddText("Begin VB.TextBox " & cControlHeader.cName)
            Case vbFrame '= 3
                cListIndex = 10
                Call AddText("Begin VB.Frame " & cControlHeader.cName)
            Case vbCommandbutton '= 4
                cListIndex = 4
                Call AddText("Begin VB.CommandButton " & cControlHeader.cName)
            Case vbCheckbox '= 5
                cListIndex = 1
                Call AddText("Begin VB.Checkbox " & cControlHeader.cName)
            Case vbOptionbutton     ' = 6
                cListIndex = 21
                Call AddText("Begin VB.Optionbutton " & cControlHeader.cName)
            Case vbCombobox     ' = 7
                cListIndex = 3
                Call AddText("Begin VB.Combobox " & cControlHeader.cName)
            Case vbListbox     '= 8
                cListIndex = 17
                Call AddText("Begin VB.ListBox " & cControlHeader.cName)
            Case vbHscroll     '= 9
                cListIndex = 12
                Call AddText("Begin VB.HScrollBar " & cControlHeader.cName)
            Case vbVscroll     '= 10
                cListIndex = 32
                Call AddText("Begin VB.VScrollBar " & cControlHeader.cName)
            Case vbTimer     '= 11
                cListIndex = 28
                Call AddText("Begin VB.Timer " & cControlHeader.cName)
            Case vbForm     '= 13
                cListIndex = 9
                Call modGlobals.LoadNewFormHolder(cControlHeader.cName)
                Call AddText("Begin VB.Form " & cControlHeader.cName)
            Case vbDriveListbox     '= 16
                cListIndex = 7
                Call AddText("Begin VB.DriveListbox " & cControlHeader.cName)
            Case vbDirectoryListbox     '= 17
                cListIndex = 6
                Call AddText("Begin VB.DirectoryListbox " & cControlHeader.cName)
            Case vbFileListbox     '= 18
                cListIndex = 8
                Call AddText("Begin VB.FileListBox " & cControlHeader.cName)
            Case vbMenu     '= 19
                cListIndex = 19
                Call AddText("Begin VB.Menu " & cControlHeader.cName)
            Case vbMDIForm     '= 20
                cListIndex = 18
                Call AddText("Begin VB.MDIForm " & cControlHeader.cName)
            Case vbShape     '= 22
                cListIndex = 26
                Call AddText("Begin VB.Shape " & cControlHeader.cName)
            Case vbLine     '= 23
                cListIndex = 16
                Call AddText("Begin VB.Line " & cControlHeader.cName)
            Case vbImage     '= 24
                cListIndex = 12
                Call AddText("Begin VB.Image " & cControlHeader.cName)
            Case vbData     '= 37
                cListIndex = 5
                Call AddText("Begin VB.Data " & cControlHeader.cName)
            Case vbOLE     '= 38
                cListIndex = 20
                Call AddText("Begin VB.OLE " & cControlHeader.cName)
            Case vbUserControl     '= 40
                cListIndex = 29
                Call AddText("Begin VB.UserControl " & cControlHeader.cName)
            Case vbPropertyPage     '= 41
                cListIndex = 24
                Call AddText("Begin VB.PropertyPage " & cControlHeader.cName)
            Case vbUserDocument     '= 42
                cListIndex = 30
                Call AddText("Begin VB.UserDocument " & cControlHeader.cName)
        End Select
        Set tliTypeInfo = tliTypeLibInfo.GetTypeInfo(Replace(Replace(lstTypeInfos.List(cListIndex), "<", ""), ">", ""))
        'Use the ItemData in lstTypeInfos to set the SearchData for lstMembers
        tliTypeLibInfo.GetMembersDirect lstTypeInfos.ItemData(cListIndex), lstMembers.hWnd, , , True
        FileLen = Loc(f) - fPos
        FileLen = cControlHeader.Length - FileLen
        
        Dim bCode As Byte 'holds gui opcode
        Dim varHold As Variant 'Holds the different data types
        Dim strHold As String 'holds the string
        Dim strReturnType As String 'holds the return type
        Do While FileLen > 1
         bCode = GetOpcode(f) 'Get the guiopcode
         
         FileLen = FileLen - 1
         Dim g As Integer
         For g = 0 To lstMembers.ListCount - 1
         
         
            'Control Postion opcode
            If bCode = 4 And cControlHeader.cType <> vbForm Then
                Dim cPosition As typeStandardControlSize
                Get f, , cPosition
                Call AddText("Left = " & cPosition.cLeft)
                Call AddText("Top = " & cPosition.cTop)
                Call AddText("Height = " & cPosition.cHeight)
                Call AddText("Width = " & cPosition.cWidth)
                FileLen = FileLen - 8
                Exit For
                'MsgBox "hey"
            End If
         
            If ReturnGuiOpcode(lstTypeInfos.ItemData(cListIndex), tliInvokeKinds, lstMembers.List(g)) = bCode Then
              'MsgBox "Prop: " & lstMembers.List(g)
                strReturnType = Trim(ReturnDataType(lstTypeInfos.ItemData(cListIndex), tliInvokeKinds, lstMembers.List(g)))
              'MsgBox "Prop: " & lstMembers.List(g) & " " & strReturnType & " Gui: " & bCode & " FileLen: " & FileLen & " Loc" & Loc(F)
                
                'Com Hack Check
                For k = 0 To UBound(gComFix)
                    If lstTypeInfos.List(cListIndex) = gComFix(k).ObjectName And lstMembers.List(g) = gComFix(k).PropertyName Then
                        strReturnType = gComFix(k).NewType
                        Exit For
                    End If
                Next
                
                If InStr(1, strReturnType, "Byte") Then
                    varHold = GetByte2(f)
                    Call AddText(lstMembers.List(g) & " = " & varHold)
                    FileLen = FileLen - 1
                    Exit For
                End If
                If InStr(1, strReturnType, "Boolean") Then
                    varHold = GetBoolean(f)
                    If varHold = True Then
                        Call AddText(lstMembers.List(g) & " = " & -1)
                    Else
                        Call AddText(lstMembers.List(g) & " = " & 0)
                    End If
                    Seek f, Loc(f)
                    FileLen = FileLen - 2
                    Exit For
                End If
                If InStr(1, strReturnType, "Integer") Then
                    varHold = GetInteger(f)
                    Call AddText(lstMembers.List(g) & " = " & varHold)
                    FileLen = FileLen - 2
                    Exit For
                End If
                If InStr(1, strReturnType, "Long") Then
                    varHold = GetLong(f)
                    Call AddText(lstMembers.List(g) & " = " & varHold)
                    FileLen = FileLen - 4
                    Exit For
                End If
                
                If InStr(1, strReturnType, "Single") Then
                    varHold = GetSingle(f)
                    Call AddText(lstMembers.List(g) & " = " & varHold)
                    FileLen = FileLen - 4
                    Exit For
                End If

                If InStr(1, strReturnType, "String") Then
                    ''Seek f, Loc(f) + 3
                    ''strHold = GetUntilNull(f)
                    strHold = GetAllString(f)
                    Call AddText(lstMembers.List(g) & " = " & Chr(34) & strHold & Chr(34))
                    FileLen = FileLen - Len(strHold) - 3
                    Exit For
                End If
                If InStr(1, strReturnType, "stdole.Picture") Then
                    
                    varHold = GetLong(f)
                    If varHold <> -1 Then
                    'MsgBox "Loc:" & Loc(f) & " " & varHold
                        Call GetStdPicture(f, varHold)
                        FileLen = FileLen - varHold - 10
                    Else
                        FileLen = FileLen - 4
                    End If
                    Exit For
                End If
               
                Exit For
            End If
            
            'Get height width top left
            If bCode = 53 Then
            '53 is the size opcode for form's
            Dim objectSize As ControlSize
                Get f, , objectSize
                FileLen = FileLen - 16
                
                If cControlHeader.cType = vbForm Then
                    Call AddText("ClientLeft = " & objectSize.clientLeft)
                    Call AddText("ClientTop = " & objectSize.clientTop)
                    Call AddText("ClientWidth = " & objectSize.clientWidth)
                    Call AddText("ClientHeight = " & objectSize.clientHeight)
                End If
                
                Exit For
            End If
         Next
        Loop
        
        
        'Get the seperator type for the end of the control
        Dim cControlEnd As Integer
        'Seek f, Loc(f) '+1
        cControlEnd = GetInteger(f)
        'MsgBox cControlEnd & " Loc:" & Loc(F)
        FileLen = file - 2
        If cControlEnd = vbFormEnd Then
            bFormEndUsed = True
            Call AddText("End")
         
        End If
        If cControlEnd = vbFormNewChildControl Then
            GoTo NewControl
        End If
        If cControlEnd = vbFormChildControl Then
            Call AddText("End")
            GoTo NewControl
        End If
        
        If cControlEnd = vbFormExistingChildControl Then
            Call AddText("End")
            GoTo NewControl
        End If
        If cControlEnd = vbFormMenu Then
            GoTo NewControl
        End If
        If bFormEndUsed = False Then
            Call AddText("End")
        End If
    End If 'For gSkipCom
    
    Next lForm 'Main Form Loop
'##########################################
'End of Form/Control Properties Loop
'##########################################
    
        'Get Object Table
        Seek f, gProjectInfo.aObjectTable + 1 - OptHeader.ImageBase
        Get #f, , gObjectTable
        
        'Resize for the number of objects...(forms,modules,classes)
        ReDim gObject(gObjectTable.ObjectCount1 - 1)
        ReDim gObjectNameArray(gObjectTable.ObjectCount1 - 1)
        
        'Get Object
        Seek f, gObjectTable.aObject + 1 - OptHeader.ImageBase
        Get #f, , gObject
        
        Dim loopC As Integer
        For loopC = 0 To UBound(gObject)
        'Get ObjectName
        Seek f, gObject(loopC).aObjectName + 1 - OptHeader.ImageBase
        gObjectNameArray(loopC) = GetUntilNull(f)
        
        
        'Get Object Info
        Seek f, gObject(loopC).aObjectInfo + 1 - OptHeader.ImageBase
        Get #f, , gObjectInfo
        
       ' MsgBox gObjectInfo.aObject + 1 - OptHeader.ImageBase
         'Get Optional Object Info
        Seek f, gObject(loopC).aObjectInfo + 57 - OptHeader.ImageBase
        Get #f, , gOptionalObjectInfo
        'Resize the control array
        If gOptionalObjectInfo.ControlCount < 5000 And gOptionalObjectInfo.ControlCount <> 0 Then
            ReDim gControl(gOptionalObjectInfo.ControlCount - 1)
            'Get Control Array
            Seek f, gOptionalObjectInfo.aControlArray + 1 - OptHeader.ImageBase
            Get #f, , gControl
            'Resize Event Table array
            ReDim gEventTable(UBound(gControl))
            'Get Event Table
            Dim i As Integer
           
            Dim ControlName As String
            For i = 0 To UBound(gControl)
               ' Seek f, gControl(i).aEventTable + 1 - OptHeader.ImageBase
               ' ReDim gEventTable(i).aEventPointer(gControl(i).EventCount - 1)
    '            Get #f, , gEventTable(i)
                If gControl(i).aName + 1 - OptHeader.ImageBase > 0 Then
                Seek f, gControl(i).aName + 1 - OptHeader.ImageBase
                ControlName = GetUntilNull(f)
                'Save the control information for the treeview
                ReDim Preserve gControlNameArray(UBound(gControlNameArray) + 1)
                gControlNameArray(UBound(gControlNameArray)).strControlName = ControlName
                gControlNameArray(UBound(gControlNameArray)).strParentForm = gObjectNameArray(loopC)
                
               ' MsgBox gControl(i).aGUID + 1 - OptHeader.ImageBase
                'MsgBox controlname  '& (gControl(i).aGUID + 1 - OptHeader.ImageBase)
                End If
            Next
        
        End If
        'Get Code info
        If gObject(loopC).ProcCount <> 0 Then
           ' MsgBox gObjectInfo.aConstantPool + 1 - OptHeader.ImageBase
            ''Seek f, gObjectInfo.aProcTable + 1 - OptHeader.ImageBase
            ''Get #f, , gCodeInfo
            If gObject(loopC).aProcNamesArray <> 0 Then
            ''Seek f, gObject(loopC).aProcNamesArray + 1 - OptHeader.ImageBase
            ''MsgBox gObject(loopC).aProcNamesArray + 1 - OptHeader.ImageBase
            End If
            'Resize the procedure array
            ReDim gProcedure(gObject(loopC).ProcCount - 1)
            'Get The Procedure array
       
            ''Seek f, gObjectInfo.aProcTable + 1 - OptHeader.ImageBase
           '' Get #f, , gProcedure
        
        End If
        Next loopC

        
    Close f

    
    'Set the compile type either pcode or ncode
    If gProjectInfo.aNativeCode <> 0 Then
        AppData.CompileType = "Native"
    Else
        AppData.CompileType = "PCode"
    End If
    
    
    MakeDir (App.Path & "\dump")
    MakeDir (App.Path & "\dump\" & FileTitle)
    mnuFileGenerate.Enabled = True
    mnuFileExportMemoryMap.Enabled = True
    
    'Get FileVersion Info
    gFileInfo = modGlobals.FileInfo(SFilePath)

    Call SetupTreeView
    Call modOutput.DumpVBExeInfo(App.Path & "\dump\" & FileTitle & "\FileReport.txt", FileTitle)
    
    'Add to recent files
    Call AddToRecentList(SFilePath, SFile)

    'Clear current data
    DDirPath = ""
    DFile = ""

  Exit Sub
    
AnalyzeError:

    MsgBox "Analyze error", vbCritical Or vbOKOnly, "Source file error"

End Sub
Sub AddToRecentList(FileName As String, FileTitle As String)
    Dim Recent1File As String
    Dim Recent1Title As String
    Dim Recent2File As String
    Dim Recent2Title As String
    Dim Recent3File As String
    Dim Recent3Title As String
    
    mnuFileSep1.Visible = True
    mnuFileRecent1.Visible = True
    
    Recent1Title = GetSetting("VB Decompiler", "Options", "Recent1FileTitle", "")
    Recent2Title = GetSetting("VB Decompiler", "Options", "Recent2FileTitle", "")
    Recent3Title = GetSetting("VB Decompiler", "Options", "Recent3FileTitle", "")
    Recent1File = GetSetting("VB Decompiler", "Options", "Recent1File", "")
    Recent2File = GetSetting("VB Decompiler", "Options", "Recent2File", "")
    Recent3File = GetSetting("VB Decompiler", "Options", "Recent3File", "")
    If Recent1Title <> "" Then
        mnuFileRecent2.Visible = True
    End If
    If Recent2Title <> "" Then
        mnuFileRecent3.Visible = True
    End If
    If Recent3Title <> "" Then
        mnuFileRecent4.Visible = True
    End If


    Call SaveSetting("VB Decompiler", "Options", "Recent4File", Recent3File)
    Call SaveSetting("VB Decompiler", "Options", "Recent4FileTitle", Recent3Title)
    Call SaveSetting("VB Decompiler", "Options", "Recent3File", Recent2File)
    Call SaveSetting("VB Decompiler", "Options", "Recent3FileTitle", Recent2Title)
    Call SaveSetting("VB Decompiler", "Options", "Recent2File", Recent1File)
    Call SaveSetting("VB Decompiler", "Options", "Recent2FileTitle", Recent1Title)


    Call SaveSetting("VB Decompiler", "Options", "Recent1File", FileName)
    Call SaveSetting("VB Decompiler", "Options", "Recent1FileTitle", FileTitle)
    
    
    
    mnuFileRecent4.Caption = mnuFileRecent3.Caption
    mnuFileRecent3.Caption = mnuFileRecent2.Caption
    mnuFileRecent2.Caption = mnuFileRecent1.Caption
    mnuFileRecent1.Caption = FileTitle
    

End Sub
Sub MakeDir(Path As String)

On Error Resume Next
    MkDir (Path)

End Sub

Private Sub mnuFileRecent1_Click()
Dim RecentTitle As String
Dim RecentFile As String
    RecentTitle = GetSetting("VB Decompiler", "Options", "Recent1FileTitle", "")
    RecentFile = GetSetting("VB Decompiler", "Options", "Recent1File", "")
    Call OpenVBExe(RecentFile, RecentTitle)
End Sub

Private Sub mnuFileRecent2_Click()
Dim RecentTitle As String
Dim RecentFile As String
    RecentTitle = GetSetting("VB Decompiler", "Options", "Recent2FileTitle", "")
    RecentFile = GetSetting("VB Decompiler", "Options", "Recent2File", "")
    Call OpenVBExe(RecentFile, RecentTitle)
End Sub

Private Sub mnuFileRecent3_Click()
Dim RecentTitle As String
Dim RecentFile As String
    RecentTitle = GetSetting("VB Decompiler", "Options", "Recent3FileTitle", "")
    RecentFile = GetSetting("VB Decompiler", "Options", "Recent3File", "")
    Call OpenVBExe(RecentFile, RecentTitle)
End Sub

Private Sub mnuFileRecent4_Click()
Dim RecentTitle As String
Dim RecentFile As String
    RecentTitle = GetSetting("VB Decompiler", "Options", "Recent4FileTitle", "")
    RecentFile = GetSetting("VB Decompiler", "Options", "Recent4File", "")
    Call OpenVBExe(RecentFile, RecentTitle)
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show vbModal, Me
    
End Sub

Private Sub tvProject_NodeClick(ByVal Node As MSComctlLib.Node)

 Dim ParentObject As Node
    Dim LenTab As Long
    Dim i As Long, o As Long
    Dim strCode As String
    
    Dim tblPath() As String
    
    If CurrentItem <> tvProject.SelectedItem.Key Then
        tblPath = Split(tvProject.SelectedItem.Key, "/")
        CurrentItem = tvProject.SelectedItem.Key

        Select Case tblPath(1)
            Case "VERSIONINFO"
                sstViewFile.TabVisible(1) = True
                sstViewFile.TabVisible(0) = False
                fxgEXEInfo.Visible = True
                fxgEXEInfo.ColAlignment(1) = 0
                fxgEXEInfo.Clear
                fxgEXEInfo.Rows = 2
                fxgEXEInfo.ColWidth(0) = 2500
                fxgEXEInfo.TextArray(0) = "Name"
                fxgEXEInfo.TextArray(1) = "Value"
                        fxgEXEInfo.ColWidth(0) = 2000
                        fxgEXEInfo.ColWidth(1) = 2500
                        fxgEXEInfo.TextArray(2) = "CompanyName"
                        fxgEXEInfo.TextArray(3) = gFileInfo.CompanyName
                        fxgEXEInfo.AddItem "FileDescription"
                        fxgEXEInfo.TextArray(5) = gFileInfo.FileDescription
                        fxgEXEInfo.AddItem "FileVersion"
                        fxgEXEInfo.TextArray(7) = gFileInfo.FileVersion
                        fxgEXEInfo.AddItem "InternalName"
                        fxgEXEInfo.TextArray(9) = gFileInfo.InternalName
                        fxgEXEInfo.AddItem "LanguageID"
                        fxgEXEInfo.TextArray(11) = gFileInfo.LanguageID
                        fxgEXEInfo.AddItem "LegalCopyright"
                        fxgEXEInfo.TextArray(13) = gFileInfo.LegalCopyright
                        fxgEXEInfo.AddItem "OrigionalFileName"
                        fxgEXEInfo.TextArray(15) = gFileInfo.OrigionalFileName
                        fxgEXEInfo.AddItem "ProductName"
                        fxgEXEInfo.TextArray(17) = gFileInfo.ProductName
                        fxgEXEInfo.AddItem "ProductVersion"
                        fxgEXEInfo.TextArray(19) = gFileInfo.ProductVersion
            Case "STRUCT"
                sstViewFile.TabVisible(1) = True
                sstViewFile.TabVisible(0) = False
                fxgEXEInfo.Visible = True
                fxgEXEInfo.ColAlignment(1) = 0
                fxgEXEInfo.Clear
                fxgEXEInfo.Rows = 2
                fxgEXEInfo.ColWidth(0) = 2500
                fxgEXEInfo.TextArray(0) = "Name"
                fxgEXEInfo.TextArray(1) = "Value"
                Select Case tblPath(2)
                    Case "", "VBHEADER"
                        fxgEXEInfo.ColWidth(0) = 2500
                        fxgEXEInfo.TextArray(2) = "Signature"
                        fxgEXEInfo.TextArray(3) = gVBHeader.Signature
                        fxgEXEInfo.AddItem "Address SubMain"
                        fxgEXEInfo.TextArray(5) = gVBHeader.aSubMain
                        fxgEXEInfo.AddItem "Address ExternalComponentTable"
                        fxgEXEInfo.TextArray(7) = gVBHeader.aExternalComponentTable
                        fxgEXEInfo.AddItem "Address GUITable"
                        fxgEXEInfo.TextArray(9) = gVBHeader.aGUITable
                        fxgEXEInfo.AddItem "Address ProjectDescription"
                        fxgEXEInfo.TextArray(11) = gVBHeader.aProjectDescription
                        fxgEXEInfo.AddItem "Address ProjectInfo"
                        fxgEXEInfo.TextArray(13) = gVBHeader.aProjectInfo
                        fxgEXEInfo.AddItem "BackupLanguageDLL"
                        fxgEXEInfo.TextArray(15) = gVBHeader.BackupLanguageDLL
                        fxgEXEInfo.AddItem "BackupLanguageID"
                        fxgEXEInfo.TextArray(17) = gVBHeader.BackupLanguageID
                        fxgEXEInfo.AddItem "Const1"
                        fxgEXEInfo.TextArray(19) = gVBHeader.Const1
                        fxgEXEInfo.AddItem "Const2"
                        fxgEXEInfo.TextArray(21) = gVBHeader.Const2
                        fxgEXEInfo.AddItem "ExternalComponentCount"
                        fxgEXEInfo.TextArray(23) = gVBHeader.ExternalComponentCount
                        fxgEXEInfo.AddItem "Flag1"
                        fxgEXEInfo.TextArray(25) = gVBHeader.Flag1
                        fxgEXEInfo.AddItem "Flag2"
                        fxgEXEInfo.TextArray(27) = gVBHeader.Flag2
                        fxgEXEInfo.AddItem "Flag3"
                        fxgEXEInfo.TextArray(29) = gVBHeader.Flag3
                        fxgEXEInfo.AddItem "Flag4"
                        fxgEXEInfo.TextArray(31) = gVBHeader.Flag4
                        fxgEXEInfo.AddItem "Flag5"
                        fxgEXEInfo.TextArray(33) = gVBHeader.Flag5
                        fxgEXEInfo.AddItem "Flag6"
                        fxgEXEInfo.TextArray(35) = gVBHeader.Flag6
                        fxgEXEInfo.AddItem "FormCount"
                        fxgEXEInfo.TextArray(37) = gVBHeader.FormCount
                        fxgEXEInfo.AddItem "LanguageDLL"
                        fxgEXEInfo.TextArray(39) = gVBHeader.LanguageDLL
                        fxgEXEInfo.AddItem "LanguageID"
                        fxgEXEInfo.TextArray(41) = gVBHeader.LanguageID
                        fxgEXEInfo.AddItem "Offset HelpFile"
                        fxgEXEInfo.TextArray(43) = gVBHeader.oHelpFile
                        fxgEXEInfo.AddItem "Offset ProjectExename"
                        fxgEXEInfo.TextArray(45) = gVBHeader.oProjectExename
                        fxgEXEInfo.AddItem "Offset ProjectName"
                        fxgEXEInfo.TextArray(47) = gVBHeader.oProjectName
                        fxgEXEInfo.AddItem "Offset ProjectTitle"
                        fxgEXEInfo.TextArray(49) = gVBHeader.oProjectTitle
                        fxgEXEInfo.AddItem "RuntimeDLLVersion"
                        fxgEXEInfo.TextArray(51) = gVBHeader.RuntimeDLLVersion
                        fxgEXEInfo.AddItem "ThreadSpace"
                        fxgEXEInfo.TextArray(53) = gVBHeader.ThreadSpace
                    Case "VBPROJECTINFO"
                        fxgEXEInfo.ColWidth(0) = 2500
                        fxgEXEInfo.TextArray(2) = "Address EndOfCode"
                        fxgEXEInfo.TextArray(3) = gProjectInfo.aEndOfCode
                        fxgEXEInfo.AddItem "Address ExternalTable"
                        fxgEXEInfo.TextArray(5) = gProjectInfo.aExternalTable
                        fxgEXEInfo.AddItem "Address NativeCode"
                        fxgEXEInfo.TextArray(7) = gProjectInfo.aNativeCode
                        fxgEXEInfo.AddItem "Address ObjectTable"
                        fxgEXEInfo.TextArray(9) = gProjectInfo.aObjectTable
                        fxgEXEInfo.AddItem "Address StartOfCode"
                        fxgEXEInfo.TextArray(11) = gProjectInfo.aStartOfCode
                        fxgEXEInfo.AddItem "Address VBAExceptionhandler"
                        fxgEXEInfo.TextArray(13) = gProjectInfo.aVBAExceptionhandler
                        fxgEXEInfo.AddItem "ExternalCount"
                        fxgEXEInfo.TextArray(15) = gProjectInfo.ExternalCount
                        fxgEXEInfo.AddItem "Flag1"
                        fxgEXEInfo.TextArray(17) = gProjectInfo.Flag1
                        fxgEXEInfo.AddItem "Flag2"
                        fxgEXEInfo.TextArray(19) = gProjectInfo.Flag2
                        fxgEXEInfo.AddItem "Flag3"
                        fxgEXEInfo.TextArray(21) = gProjectInfo.Flag3
                        fxgEXEInfo.AddItem "Null1"
                        fxgEXEInfo.TextArray(23) = gProjectInfo.Null1
                        fxgEXEInfo.AddItem "NullSpacer"
                        fxgEXEInfo.TextArray(25) = gProjectInfo.NullSpacer
                        fxgEXEInfo.AddItem "oProjectLocation"
                        fxgEXEInfo.TextArray(27) = gProjectInfo.oProjectLocation
                        fxgEXEInfo.AddItem "OriginalPathName"
                        fxgEXEInfo.TextArray(29) = gProjectInfo.OriginalPathName
                        fxgEXEInfo.AddItem "Signature"
                        fxgEXEInfo.TextArray(31) = gProjectInfo.Signature
                        fxgEXEInfo.AddItem "ThreadSpace"
                        fxgEXEInfo.TextArray(33) = gProjectInfo.ThreadSpace
                    Case "VBOBJECTABLE"
                        fxgEXEInfo.ColWidth(0) = 2500
                        fxgEXEInfo.TextArray(2) = "Address 1"
                        fxgEXEInfo.TextArray(3) = gObjectTable.Address1
                        fxgEXEInfo.AddItem "Address 2"
                        fxgEXEInfo.TextArray(5) = gObjectTable.Address2
                        fxgEXEInfo.AddItem "Address 3"
                        fxgEXEInfo.TextArray(7) = gObjectTable.Address3
                        fxgEXEInfo.AddItem "Address of First Object"
                        fxgEXEInfo.TextArray(9) = gObjectTable.aObject
                        fxgEXEInfo.AddItem "Address of ProjectName"
                        fxgEXEInfo.TextArray(11) = gObjectTable.aProjectName
                        fxgEXEInfo.AddItem "Const1"
                        fxgEXEInfo.TextArray(13) = gObjectTable.Const1
                        fxgEXEInfo.AddItem "Const2"
                        fxgEXEInfo.TextArray(15) = gObjectTable.Const2
                        fxgEXEInfo.AddItem "Const3"
                        fxgEXEInfo.TextArray(17) = gObjectTable.Const3
                        fxgEXEInfo.AddItem "Flag1"
                        fxgEXEInfo.TextArray(19) = gObjectTable.Flag1
                        fxgEXEInfo.AddItem "Flag2"
                        fxgEXEInfo.TextArray(21) = gObjectTable.Flag2
                        fxgEXEInfo.AddItem "Flag3"
                        fxgEXEInfo.TextArray(23) = gObjectTable.Flag3
                        fxgEXEInfo.AddItem "Flag4"
                        fxgEXEInfo.TextArray(25) = gObjectTable.Flag4
                        fxgEXEInfo.AddItem "LangID1"
                        fxgEXEInfo.TextArray(27) = gObjectTable.LangID1
                        fxgEXEInfo.AddItem "LangID2"
                        fxgEXEInfo.TextArray(29) = gObjectTable.LangID2
                        fxgEXEInfo.AddItem "Null1"
                        fxgEXEInfo.TextArray(31) = gObjectTable.Null1
                        fxgEXEInfo.AddItem "Null2"
                        fxgEXEInfo.TextArray(33) = gObjectTable.Null2
                        fxgEXEInfo.AddItem "Null3"
                        fxgEXEInfo.TextArray(35) = gObjectTable.Null3
                        fxgEXEInfo.AddItem "Null4"
                        fxgEXEInfo.TextArray(37) = gObjectTable.Null4
                        fxgEXEInfo.AddItem "Null5"
                        fxgEXEInfo.TextArray(39) = gObjectTable.Null5
                        fxgEXEInfo.AddItem "Null6"
                        fxgEXEInfo.TextArray(41) = gObjectTable.Null6
                        fxgEXEInfo.AddItem "ObjectCount1"
                        fxgEXEInfo.TextArray(43) = gObjectTable.ObjectCount1
                        fxgEXEInfo.AddItem "ObjectCount2"
                        fxgEXEInfo.TextArray(45) = gObjectTable.ObjectCount2
                        fxgEXEInfo.AddItem "ObjectCount3"
                        fxgEXEInfo.TextArray(47) = gObjectTable.ObjectCount3


                End Select
            Case "EXEDATA"  '#####################################################'
                sstViewFile.TabVisible(1) = True
                sstViewFile.TabVisible(0) = False
                fxgEXEInfo.Visible = True
                
                fxgEXEInfo.ColAlignment(1) = 0
                fxgEXEInfo.Clear
                fxgEXEInfo.Rows = 2
                fxgEXEInfo.ColWidth(0) = 2000
                fxgEXEInfo.TextArray(0) = "Name"
                fxgEXEInfo.TextArray(1) = "Value"
                
                Select Case tblPath(2)
                    Case "", "EXEHEADER"
                        fxgEXEInfo.ColWidth(0) = 1500
                        fxgEXEInfo.TextArray(2) = "Signature"
                        fxgEXEInfo.TextArray(3) = DosHeader.Magic
                        fxgEXEInfo.AddItem "Extra Bytes"
                        fxgEXEInfo.TextArray(5) = DosHeader.NumBytesLastPage
                        fxgEXEInfo.AddItem "Pages"
                        fxgEXEInfo.TextArray(7) = DosHeader.NumPages
                        fxgEXEInfo.AddItem "Reloc Items"
                        fxgEXEInfo.TextArray(9) = DosHeader.NumRelocates
                        fxgEXEInfo.AddItem "Header Size"
                        fxgEXEInfo.TextArray(11) = DosHeader.NumHeaderBlks
                        fxgEXEInfo.AddItem "Min Alloc"
                        fxgEXEInfo.TextArray(13) = DosHeader.ReservedW8
                        fxgEXEInfo.AddItem "Max Alloc"
                        fxgEXEInfo.TextArray(15) = DosHeader.ReservedW9
                        fxgEXEInfo.AddItem "Initial SS"
                        fxgEXEInfo.TextArray(17) = DosHeader.SSPointer
                        fxgEXEInfo.AddItem "Initial SP"
                        fxgEXEInfo.TextArray(19) = DosHeader.SPPointer
                        fxgEXEInfo.AddItem "Check Sum"
                        fxgEXEInfo.TextArray(21) = DosHeader.Checksum
                        fxgEXEInfo.AddItem "Initial IP"
                        fxgEXEInfo.TextArray(23) = DosHeader.IPPointer
                        fxgEXEInfo.AddItem "Initial CS"
                        fxgEXEInfo.TextArray(25) = DosHeader.CurrentSeg
                        fxgEXEInfo.AddItem "Reloc Table"
                        fxgEXEInfo.TextArray(27) = DosHeader.RelocTablePointer
                        fxgEXEInfo.AddItem "Overlay"
                        fxgEXEInfo.TextArray(29) = DosHeader.Overlay
                    Case "COFFHEADER"
                        fxgEXEInfo.ColWidth(0) = 2000
                        fxgEXEInfo.TextArray(2) = "Signature"
                        fxgEXEInfo.TextArray(3) = PeHeader.Magic
                        fxgEXEInfo.AddItem "Machine"
                        fxgEXEInfo.TextArray(5) = PeHeader.Machine
                        fxgEXEInfo.AddItem "Number Of Sections"
                        fxgEXEInfo.TextArray(7) = PeHeader.NumSections
                        fxgEXEInfo.AddItem "Time Date Stamp"
                        fxgEXEInfo.TextArray(9) = PeHeader.TimeDate
                        fxgEXEInfo.AddItem "Pointer To Symbol Table"
                        fxgEXEInfo.TextArray(11) = PeHeader.SymbolTablePointer
                        fxgEXEInfo.AddItem "Number Of Symbols"
                        fxgEXEInfo.TextArray(13) = PeHeader.NumSymbols
                        fxgEXEInfo.AddItem "Optional Header Size"
                        fxgEXEInfo.TextArray(15) = PeHeader.OptionalHdrSize
                        fxgEXEInfo.AddItem "Characteristics"
                        fxgEXEInfo.TextArray(17) = PeHeader.Properties
                    Case "OPTIONALHEADER"
                        
                        fxgEXEInfo.ColWidth(0) = 2500
                        fxgEXEInfo.TextArray(2) = "Magic"
                        fxgEXEInfo.TextArray(3) = modPeSkeleton.OptHeader.Magic
                        fxgEXEInfo.AddItem "Linker Major Version"
                        fxgEXEInfo.TextArray(5) = modPeSkeleton.OptHeader.MajLinkerVer
                        fxgEXEInfo.AddItem "Linker Minor Version"
                        fxgEXEInfo.TextArray(7) = modPeSkeleton.OptHeader.MinLinkerVer
                        fxgEXEInfo.AddItem "Size Of Code Section"
                        fxgEXEInfo.TextArray(9) = modPeSkeleton.OptHeader.CodeSize
                        fxgEXEInfo.AddItem "Initialized DataSize"
                        fxgEXEInfo.TextArray(11) = modPeSkeleton.OptHeader.InitDataSize
                        fxgEXEInfo.AddItem "Uninitialized DataSize"
                        fxgEXEInfo.TextArray(13) = modPeSkeleton.OptHeader.UninitDataSize
                        fxgEXEInfo.AddItem "Entry Point RVA"
                        fxgEXEInfo.TextArray(15) = modPeSkeleton.OptHeader.EntryPoint
                        fxgEXEInfo.AddItem "Base Of Code"
                        fxgEXEInfo.TextArray(17) = modPeSkeleton.OptHeader.CodeBase
                        fxgEXEInfo.AddItem "Base Of Data"
                        fxgEXEInfo.TextArray(19) = modPeSkeleton.OptHeader.DataBase
                        fxgEXEInfo.AddItem "Image Base"
                        fxgEXEInfo.TextArray(21) = modPeSkeleton.OptHeader.ImageBase
                        fxgEXEInfo.AddItem "Section Alignement"
                        fxgEXEInfo.TextArray(23) = modPeSkeleton.OptHeader.SectionAlignment
                        fxgEXEInfo.AddItem "File Alignement"
                        fxgEXEInfo.TextArray(25) = modPeSkeleton.OptHeader.FileAlignment
                        fxgEXEInfo.AddItem "OS Major Version"
                        fxgEXEInfo.TextArray(27) = modPeSkeleton.OptHeader.MajOSVer
                        fxgEXEInfo.AddItem "OS Minor Version"
                        fxgEXEInfo.TextArray(29) = modPeSkeleton.OptHeader.MinOSVer
                        fxgEXEInfo.AddItem "User Major Version" 'bad
                        fxgEXEInfo.TextArray(31) = modPeSkeleton.OptHeader.MajImageVer
                        fxgEXEInfo.AddItem "User Minor Version" 'bad
                        fxgEXEInfo.TextArray(33) = modPeSkeleton.OptHeader.MinImageVer
                        fxgEXEInfo.AddItem "Sub Sys Major Version"
                        fxgEXEInfo.TextArray(35) = modPeSkeleton.OptHeader.MajSSysVer
                        fxgEXEInfo.AddItem "Sub Sys Minor Version"
                        fxgEXEInfo.TextArray(37) = modPeSkeleton.OptHeader.MinSSysVer
                        fxgEXEInfo.AddItem "Reserved" 'bad
                        fxgEXEInfo.TextArray(39) = modPeSkeleton.OptHeader.SSizeRes
                        fxgEXEInfo.AddItem "Image Size"
                        fxgEXEInfo.TextArray(41) = modPeSkeleton.OptHeader.SizeImage
                        fxgEXEInfo.AddItem "Header Size"
                        fxgEXEInfo.TextArray(43) = modPeSkeleton.OptHeader.SizeHeader
                        fxgEXEInfo.AddItem "File Checksum"
                        fxgEXEInfo.TextArray(45) = modPeSkeleton.OptHeader.Checksum
                        fxgEXEInfo.AddItem "Sub System"
                        fxgEXEInfo.TextArray(47) = modPeSkeleton.OptHeader.SSystem
                        fxgEXEInfo.AddItem "DLL Flags" 'bad
                        fxgEXEInfo.TextArray(49) = modPeSkeleton.OptHeader.LFlags
                        fxgEXEInfo.AddItem "Stack Reserved Size"
                        fxgEXEInfo.TextArray(51) = modPeSkeleton.OptHeader.SSizeRes
                        fxgEXEInfo.AddItem "Stack Commit Size"
                        fxgEXEInfo.TextArray(53) = modPeSkeleton.OptHeader.SSizeCom
                        fxgEXEInfo.AddItem "Heap Reserved Size"
                        fxgEXEInfo.TextArray(55) = modPeSkeleton.OptHeader.HSizeRes
                        fxgEXEInfo.AddItem "Heap Commit Size"
                        fxgEXEInfo.TextArray(57) = modPeSkeleton.OptHeader.HSizeCom
                        fxgEXEInfo.AddItem "Loader Flags"
                        fxgEXEInfo.TextArray(59) = modPeSkeleton.OptHeader.LFlags
                    Case "SECTIONHEADER"
                        If tblPath(3) <> "" Then
                            Dim SelSection As Long
                            SelSection = Val(tblPath(3))
                            
                            fxgEXEInfo.ColWidth(0) = 2000
                            fxgEXEInfo.TextArray(2) = "Section Name"
                            fxgEXEInfo.TextArray(3) = modPeSkeleton.SecHeader(SelSection).SecName
                           
                            fxgEXEInfo.AddItem "Virtual Size"
                            fxgEXEInfo.TextArray(5) = modPeSkeleton.SecHeader(SelSection).Properties
                            fxgEXEInfo.AddItem "RVA Offset"
                            fxgEXEInfo.TextArray(7) = modPeSkeleton.SecHeader(SelSection).Address
                            fxgEXEInfo.AddItem "Size Of Raw Data"
                            fxgEXEInfo.TextArray(9) = modPeSkeleton.SecHeader(SelSection).SizeRawData
                            fxgEXEInfo.AddItem "Pointer To Raw Data"
                            fxgEXEInfo.TextArray(11) = modPeSkeleton.SecHeader(SelSection).RawDataPointer
                            fxgEXEInfo.AddItem "Pointer To Relocs"
                            fxgEXEInfo.TextArray(13) = modPeSkeleton.SecHeader(SelSection).RelocationPointer
                            fxgEXEInfo.AddItem "Pointer To Line Numbers"
                            fxgEXEInfo.TextArray(15) = modPeSkeleton.SecHeader(SelSection).LineNumPointer
                            fxgEXEInfo.AddItem "Number Of Relocs"
                            fxgEXEInfo.TextArray(17) = modPeSkeleton.SecHeader(SelSection).NumRelocations
                            fxgEXEInfo.AddItem "Number Of Line Numbers"
                            fxgEXEInfo.TextArray(19) = modPeSkeleton.SecHeader(SelSection).NumLineNumbers
                            fxgEXEInfo.AddItem "Section Flags"
                            fxgEXEInfo.TextArray(21) = modPeSkeleton.SecHeader(SelSection).Misc
                        Else
                        'SecHeader(0).Address
                            fxgEXEInfo.TextArray(2) = " " & ExtString(SecHeader(0).SecName)
                            fxgEXEInfo.TextArray(3) = AddChar(Hex(SecHeader(0).Address), 8)
                           For i = 1 To PeHeader.NumSections
                               fxgEXEInfo.AddItem " " & ExtString(SecHeader(0).SecName)
                               fxgEXEInfo.TextArray(3 + i * 2) = AddChar(Hex(SecHeader(0).Address), 8)
                            Next i
                        End If
                End Select
            Case "PROJECT"  '#####################################################'
                sstViewFile.TabVisible(0) = True
                sstViewFile.TabVisible(1) = False
                fxgEXEInfo.Visible = False
                Call modOutput.ShowVBPFile

            Case "CODE"     '#####################################################'
                sstViewFile.TabVisible(0) = True
                sstViewFile.TabVisible(1) = False
                fxgEXEInfo.Visible = False
            Case "FORMS"
                If tblPath(2) <> "" Then
                    sstViewFile.TabVisible(0) = True
                    sstViewFile.TabVisible(1) = False
                    fxgEXEInfo.Visible = False
                    
                    For i = 0 To txtFinal.UBound
                        If UCase(txtFinal(i).Tag) = tblPath(2) Then
                            txtCode.Text = txtFinal(i).Text
                            Exit For
                        End If
                   
                    Next
                End If

        End Select
    
     
    End If
End Sub
Private Sub SetupTreeView()
'Sets up all the nodes in the Treeview control
    Dim Parent(0 To &HFF) As String, LenTab As Long, IsMenu As Boolean
    Dim i As Long, o As Long, e As Long
    FileName = SFile
    Call tvProject.Nodes.Add(, , "ROOT/PROJECT/" & FileName, Mid(FileName, InStrRev(FileName, "\") + 1), 34)

    tvProject.Nodes(1).Selected = True
    tvProject.Nodes(1).Expanded = True
    tvProject_NodeClick tvProject.Nodes(1)
    
    '####################   Information about the exe  ####################'
    Call tvProject.Nodes.Add("ROOT/PROJECT/" & FileName, tvwChild, "ROOT/EXEDATA/", "PE Header", 1)
    Parent(0) = "ROOT/EXEDATA/"
    tvProject.Nodes.Add Parent(0), tvwChild, "ROOT/EXEDATA/EXEHEADER/", "EXE Header", 2
    tvProject.Nodes.Add Parent(0), tvwChild, "ROOT/EXEDATA/COFFHEADER/", "Coff Header", 2
    tvProject.Nodes.Add Parent(0), tvwChild, "ROOT/EXEDATA/OPTIONALHEADER/", "Optional Header", 2
    tvProject.Nodes.Add Parent(0), tvwChild, "ROOT/EXEDATA/SECTIONHEADER/", "Section Header", 3
    
   For i = 1 To PeHeader.NumSections
     tvProject.Nodes.Add "ROOT/EXEDATA/SECTIONHEADER/", tvwChild, "ROOT/EXEDATA/SECTIONHEADER/" & i & "/", SecHeader(i).SecName, 2
    Next i
    '####################   VB Strutures       ####################'
    Call tvProject.Nodes.Add("ROOT/PROJECT/" & FileName, tvwChild, "ROOT/STRUCT/", "VB Structures", 1)
    Parent(0) = "ROOT/STRUCT/"
    tvProject.Nodes.Add Parent(0), tvwChild, "ROOT/STRUCT/VBHEADER/", "VB Header", 2
    tvProject.Nodes.Add Parent(0), tvwChild, "ROOT/STRUCT/VBPROJECTINFO/", "VB Project Information", 2
    tvProject.Nodes.Add Parent(0), tvwChild, "ROOT/STRUCT/VBOBJECTABLE/", "VB Object Table", 2
    tvProject.Nodes.Add Parent(0), tvwChild, "ROOT/STRUCT/VBEVENTTABLE/", "VB Event Table", 2
    tvProject.Nodes.Add Parent(0), tvwChild, "ROOT/STRUCT/VBEXTERNALTABLE/", "VB External Table", 2
    '####################   VB Forms       ####################'
    Call tvProject.Nodes.Add("ROOT/PROJECT/" & FileName, tvwChild, "ROOT/FORMS/", "Forms", 1)
    Parent(0) = "ROOT/FORMS/"
    For i = 0 To UBound(gObject)
        If gObject(i).ObjectType = 98435 Then
            tvProject.Nodes.Add Parent(0), tvwChild, "ROOT/FORMS/" & UCase(gObjectNameArray(i)) & "/", gObjectNameArray(i), 10
        End If
    Next
    For i = 1 To UBound(gControlNameArray)
        If gControlNameArray(i).strControlName <> "" And gControlNameArray(i).strControlName <> "Form" Then
            On Error Resume Next
            tvProject.Nodes.Add "ROOT/FORMS/" & UCase(gControlNameArray(i).strParentForm) & "/", tvwChild, "ROOT/FORMS/" & UCase(gControlNameArray(i).strParentForm) & "/" & i & "/", gControlNameArray(i).strControlName, 2
        End If
    Next
    '####################   VB Modules       ####################'
    Call tvProject.Nodes.Add("ROOT/PROJECT/" & FileName, tvwChild, "ROOT/MODS/", "Modules", 1)
    Parent(0) = "ROOT/MODS/"
    AppData.AppModuleCount = 0
    For i = 0 To UBound(gObject)
        If gObject(i).ObjectType = 98305 Then
            tvProject.Nodes.Add Parent(0), tvwChild, "ROOT/MODS/" & UCase(gObjectNameArray(i)) & "/", gObjectNameArray(i), 40
            AppData.AppModuleCount = AppData.AppModuleCount + 1
        End If
    Next
    '####################   VB Classes       ####################'
    Call tvProject.Nodes.Add("ROOT/PROJECT/" & FileName, tvwChild, "ROOT/CLASS/", "Classes", 1)
    Parent(0) = "ROOT/CLASS/"
    For i = 0 To UBound(gObject)
        If gObject(i).ObjectType = 1146883 Then
            tvProject.Nodes.Add Parent(0), tvwChild, "ROOT/CLASS/" & UCase(gObjectNameArray(i)) & "/", gObjectNameArray(i), 41
        End If
    Next
    '####################   Procedures - Code       ####################'

    Call tvProject.Nodes.Add("ROOT/PROJECT/" & FileName, tvwChild, "ROOT/CODE/", "Procedures - Code", 1)
    Parent(0) = "ROOT/CODE/"
    
    tvProject.Nodes.Add(Parent(0), tvwChild, "ROOT/CODE/" & "ASM", "Code assembly", 4).Tag = -2
    '####################   Images     ####################'
    Call tvProject.Nodes.Add("ROOT/PROJECT/" & FileName, tvwChild, "ROOT/IMAEGS/", "Images", 1)
    '####################   File Version Information    ####################'
    Call tvProject.Nodes.Add("ROOT/PROJECT/" & FileName, tvwChild, "ROOT/VERSIONINFO/", "File Version Information", 1)



    CurrentNode = 0

    tvProject_NodeClick tvProject.Nodes(1)
End Sub
Function GetOpcode(FileNum As Variant) As Byte
'Function just retrieves a byte used to get gui opcode
    Dim opcode As Byte
    Get FileNum, , opcode
    GetOpcode = opcode
End Function
Private Function DumpObject(FileNum As Variant, ObjectName As String, Length As Integer, FileStart As Long, HeaderEnd As Long) As Long
'Dumps a Gui Object
  On Error GoTo bad
    MakeDir (App.Path & "\dump")
    MakeDir (App.Path & "\dump\" & SFile)
    Dim bArray() As Byte
    ReDim bArray(Length)
    'Get the ojbect information
    Seek FileNum, FileStart + 1
    
    Get FileNum, , bArray
    Dim fFileEnd As Long
    fFileEnd = Loc(FileNum)
    Seek FileNum, HeaderEnd
    'Save the information
    Open App.Path & "\dump\" & SFile & "\" & ObjectName & ".txt" For Binary Access Write Lock Write As #12
        Put #12, , bArray
    Close #12
    
    DumpObject = (fFileEnd + 1)
    Exit Function
bad:
    DumpObject = -1
Exit Function
End Function
Sub GetStdPicture(FileNum As Variant, Length As Variant)
   ' MsgBox "STDPICTURE!!!!" & " Loc:" & Loc(FileNum)
    Dim picHeader As typePictureHeader
    'Get Picture Header
    Get FileNum, , picHeader
    Length = Length - 8
    'Get the whole picture
    Dim i As Integer
    Dim bData As Byte
   ' MsgBox Length & " Loc" & Loc(FileNum)
    Open App.Path & "\test2.ico" For Binary Access Write Lock Write As #23
    
        For i = 1 To Length
            Get FileNum, , bData
            Put #23, , bData
        Next
    
    Close #23

End Sub
