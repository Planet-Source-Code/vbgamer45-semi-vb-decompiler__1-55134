VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Tag             =   "                                  vbgamer45"
   Begin VB.CheckBox chkDumpControls 
      Caption         =   "Dump Control/Form raw binary data"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   3015
   End
   Begin VB.CheckBox chkSkipCOM 
      Caption         =   "Skip COM and Control/Form Property Processing"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   4335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   2520
      Width           =   1815
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkDumpControls_Click()
    If chkDumpControls.Value = vbChecked Then
        gDumpData = True
    Else
        gDumpData = False
    End If
End Sub

Private Sub chkSkipCOM_Click()
    If chkSkipCOM.Value = vbChecked Then
        gSkipCom = True
    Else
        gSkipCom = False
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If gSkipCom = True Then chkSkipCOM.Value = vbChecked
    If gDumpData = True Then Me.chkDumpControls.Value = vbChecked
End Sub
