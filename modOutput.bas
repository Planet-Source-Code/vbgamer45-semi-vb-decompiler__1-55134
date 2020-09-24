Attribute VB_Name = "modOutput"
Sub DumpVBExeInfo(FileName As String, FileTitle As String)
'Prints  a report about the Exe that was decompiled
Dim i As Integer

    Open FileName For Output As #1
        Print #1, "----------------------------------"
        Print #1, FileTitle
        Print #1, "Output made by Semi VB Decompiler by vbgamer45"
        Print #1, "----------------------------------"
        Print #1, "VB Exe Info"
        Print #1, "----------------------------------"
        Print #1, "VBStartOffset " & AppData.VBStartOffset
        Print #1, "FormCount= " & gVBHeader.FormCount
        Print #1, "ModuleCount= " & AppData.AppModuleCount
        Print #1, "CompileType= " & AppData.CompileType
        Print #1, "----------------------------------"
        Print #1, "VB Header Infomation"
        Print #1, "----------------------------------"
        Print #1, "ProjectTitle= " & ProjectTitle
        Print #1, "ProjectName= " & ProjectName
        Print #1, "ExeName= " & ProjectExename
        Print #1, "HelpFile= " & HelpFile
        If gVBHeader.aSubMain <> 0 Then
            Print #1, "SubMain Address= " & gVBHeader.aSubMain + 1 - OptHeader.ImageBase
        End If
        Print #1, "ExternalComponentCount= " & gVBHeader.ExternalComponentCount
        Print #1, "----------------------------------"
        Print #1, "Object List"
        Print #1, "----------------------------------"
        For i = 0 To UBound(gObjectNameArray)
            Print #1, gObjectNameArray(i)
        Next i
    Close #1
    
    
End Sub
Sub WriteVBP(FileName As String)
'Writes the visual basic project file
    Dim DATAHERE As String
    Open FileName For Output As #3
    
        Print #3, "Type=Exe"
        
        Print #3, "Reference="
        Print #3, "Object="
        
        
        For i = 0 To UBound(gObject)
            If gObject(i).ObjectType = 98435 Then
            'Form LooP
                Print #3, "Form=" & gObjectNameArray(i) & ".frm"
            End If
            If gObject(i).ObjectType = 98305 Then
            'Module Loop
                Print #3, "Module=" & gObjectNameArray(i) & "; " & gObjectNameArray(i) & ".bas"
            End If
            If gObject(i).ObjectType = 1146883 Then
            'Class Loop
                Print #3, "Class=" & gObjectNameArray(i) & "; " & gObjectNameArray(i) & ".cls"
            End If
        Next
        
        
        Print #3, "IconForm=" & Chr(34) & DATAHERE & Chr(34)
        If gVBHeader.aSubMain = 0 Then
            Print #3, "Startup=" & Chr(34) & gObjectNameArray(0) & Chr(34)
        End If
        Print #3, "Description=" & Chr(34) & ProjectDescription & Chr(34)
        Print #3, "HelpFile=" & Chr(34) & HelpFile & Chr(34)
        Print #3, "Name=" & Chr(34) & ProjectName & Chr(34)
        Print #3, "Title=" & Chr(34) & ProjectTitle & Chr(34)
        Print #3, "ExeName32=" & Chr(34) & ProjectExename & Chr(34)
        Print #3, "VersionCompanyName=" & Chr(34) & gFileInfo.CompanyName & Chr(34)
    Close #3
    
End Sub
Sub ShowVBPFile()
    frmMain.txtCode.Text = ""
    frmMain.txtCode.Text = frmMain.txtCode.Text & "Type=Exe" & vbCrLf
        For i = 0 To UBound(gObject)
            If gObject(i).ObjectType = 98435 Then
            'Form LooP
                frmMain.txtCode.Text = frmMain.txtCode.Text & "Form=" & gObjectNameArray(i) & ".frm" & vbCrLf
            End If
            If gObject(i).ObjectType = 98305 Then
            'Module Loop
                frmMain.txtCode.Text = frmMain.txtCode.Text & "Module=" & gObjectNameArray(i) & "; " & gObjectNameArray(i) & ".bas" & vbCrLf
            End If
            If gObject(i).ObjectType = 1146883 Then
            'Class Loop
                frmMain.txtCode.Text = frmMain.txtCode.Text & "Class=" & gObjectNameArray(i) & "; " & gObjectNameArray(i) & ".cls" & vbCrLf
            End If
        Next
    If gVBHeader.aSubMain = 0 Then
        frmMain.txtCode.Text = frmMain.txtCode.Text & "Startup=" & Chr(34) & gObjectNameArray(0) & Chr(34) & vbCrLf
    End If
    frmMain.txtCode.Text = frmMain.txtCode.Text & "Description=" & Chr(34) & ProjectDescription & Chr(34) & vbCrLf
    frmMain.txtCode.Text = frmMain.txtCode.Text & "HelpFile=" & Chr(34) & HelpFile & Chr(34) & vbCrLf
    frmMain.txtCode.Text = frmMain.txtCode.Text & "Name=" & Chr(34) & ProjectName & Chr(34) & vbCrLf
    frmMain.txtCode.Text = frmMain.txtCode.Text & "Title=" & Chr(34) & ProjectTitle & Chr(34) & vbCrLf
    frmMain.txtCode.Text = frmMain.txtCode.Text & "ExeName32=" & Chr(34) & ProjectExename & Chr(34) & vbCrLf
    frmMain.txtCode.Text = frmMain.txtCode.Text & "VersionCompanyName=" & Chr(34) & gFileInfo.CompanyName & Chr(34) & vbCrLf

End Sub
Sub WriteForms(FilePath As String)
    For i = 0 To frmMain.txtFinal.UBound
        If frmMain.txtFinal(i).Tag <> "" Then
        Open FilePath & frmMain.txtFinal(i).Tag & ".frm" For Output As #4
            Print #4, "VERSION 5.00"
            'Begin Object References
            
            'Begin Form
            Print #4, frmMain.txtFinal(i).Text
        Close #4
        End If
    Next
    
End Sub
Sub WriteFromFrx()
'Write the forms graphic files

End Sub
Sub WriteModules(FileName As String, ObjectName As String)
Dim DATAHERE As String
    Open FileName For Output As #5
        Print #5, "Attribute VB_Name = " & Chr(34) & ObjectName & Chr(34)
    Close #5
End Sub
Sub WriteClasses(FileName As String, ObjectName As String)
Dim DATAHERE As String
    Open FileName For Output As #6
        Print #6, "VERSION 1.0 CLASS"
        Print #6, "Begin"
        Print #6, "  MultiUse = -1  'True"
        Print #6, "  Persistable = 0  'NotPersistable"
        Print #6, "  DataBindingBehavior = 0  'vbNone"
        Print #6, "  DataSourceBehavior = 0   'vbNone"
        Print #6, "  MTSTransactionMode = 0   'NotAnMTSObject"
        Print #6, "End"
        Print #6, "Attribute VB_Name = " & Chr(34) & ObjectName & Chr(34)
        Print #6, "Attribute VB_GlobalNameSpace = False"
        Print #6, "Attribute VB_Creatable = True"
        Print #6, "Attribute VB_PredeclaredId = False"
        Print #6, "Attribute VB_Exposed = False"
        Print #6, "Attribute VB_Ext_KEY = " & Chr(34) & "SavedWithClassBuilder6" & Chr(34) & "," & Chr(34) & "Yes" & Chr(34)
        Print #6, "Attribute VB_Ext_KEY = " & Chr(34) & "Top_Level" & Chr(34) & " ," & Chr(34) & "No" & Chr(34)
    Close #6
End Sub
Sub WriteUserControls(FileName As String)
Dim DATAHERE As String
    Open FileName For Output As #7
    
    Close #7
End Sub
