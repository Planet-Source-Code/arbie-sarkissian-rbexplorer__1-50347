Attribute VB_Name = "mdul_Main"
Option Explicit

Public v_iIndexCounter As Integer
Public a_iBrowserIndex() As Integer
Public a_sAddress() As String
Public a_lProgress() As Long
Public a_lProgressMAX() As Long
Public v_iActiveProgressIndex As Integer
Public v_sCurrentNavigationAddress As String

Public v_rsBookmarks As New Recordset
Public v_rsHistory As New Recordset

Public v_iBackIndex As Integer
Public a_sBack() As String
Public v_iForwardIndex As Integer
Public a_sForward() As String

Public v_iPortNumber As Integer
Public v_sServerID As String
Public v_sClientID As String
Public v_sRemoteIP As String
Public v_bAutomaticClientConnect As Boolean

Public v_iCoolbarHeight As Integer

Public v_iBookmarksIndex As Integer

Public v_bDownloadRequested As Boolean
Public v_sDownloadFileName As String
Public v_lDownloadFileLen As Long
Public v_sDownloadedData As String
Public v_sTotalDownloadedData As String
Public v_lDownloadedBytes As Long

Public v_sConnectionString As String

Sub Main()
    Dim v_rsData As New Recordset

    With frm_Main
        v_sConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database.mdb"
    
        .cbo_Address.ComboItems.Add 1, , "", 1
        .cbo_Address.ComboItems(1).Selected = True
        
        ReDim Preserve a_iBrowserIndex(0)
        a_iBrowserIndex(0) = 0
        
        ReDim Preserve a_lProgress(0)
        ReDim Preserve a_lProgressMAX(0)
        
        ReDim Preserve a_sAddress(0)
        
        ReDim Preserve a_sBack(5)
        ReDim Preserve a_sForward(5)
                        
        .sst_General.Top = 480
        .sst_General.TabPicture(0) = frm_Main.pic_Bookmarks.Image
        .sst_General.TabPicture(1) = frm_Main.pic_History.Image
        .sst_General.TabPicture(2) = frm_Main.pic_Source.Image
        .sst_General.TabPicture(3) = frm_Main.pic_Client.Image
        .sst_General.TabPicture(4) = frm_Main.pic_Server.Image
        
        .cbo_Address.ComboItems(1).Selected = True
        
        v_sServerID = "Server"
        v_iPortNumber = 10011
        .sok_Server.LocalPort = v_iPortNumber
        .sok_Server.Listen
        
        v_iCoolbarHeight = 435
        .cbo_SearchList.ComboItems.Add , , "Google", 3
        .cbo_SearchList.ComboItems.Add , , "Yahoo!", 6
        .cbo_SearchList.ComboItems.Add , , "MSN", 7
        .cbo_SearchList.ComboItems.Add , , "Goto", 5
        .cbo_SearchList.ComboItems.Add , , "Excite", 4
        .cbo_SearchList.ComboItems.Add , , "AltaVista", 3
        
        v_rsData.Open "SELECT * FROM Bookmarks", v_sConnectionString
        While Not v_rsData.EOF
            If v_iBookmarksIndex <> 0 Then Load frm_Main.pdi_Bookmarks(v_iBookmarksIndex)
            frm_Main.pdi_Bookmarks(v_iBookmarksIndex).Caption = v_rsData.Fields(2).Value
            v_iBookmarksIndex = v_iBookmarksIndex + 1
            v_rsData.MoveNext
        Wend
        v_rsData.Close
        Set v_rsData = Nothing
        
        SetIcon .hWnd, "icn", False
        .Show
    End With
End Sub

Public Sub ShowGeneralTab(m_Value As Boolean)
    Dim v_iLoop As Integer
    
    If m_Value = True Then
        For v_iLoop = 0 To v_iIndexCounter
            frm_Main.web_Browser(a_iBrowserIndex(v_iLoop)).Left = frm_Main.sst_General.Width + 45
            If frm_Main.sst_General.Visible = False Then
                frm_Main.web_Browser(a_iBrowserIndex(v_iLoop)).Width = frm_Main.web_Browser(a_iBrowserIndex(v_iLoop)).Width - frm_Main.sst_General.Width - 45
            End If
            frm_Main.tab_Main.Left = frm_Main.sst_General.Width + 45
            frm_Main.sst_General.Visible = True
        Next v_iLoop
    Else
        For v_iLoop = 0 To v_iIndexCounter
            frm_Main.web_Browser(a_iBrowserIndex(v_iLoop)).Left = 0
            frm_Main.web_Browser(a_iBrowserIndex(v_iLoop)).Width = frm_Main.web_Browser(a_iBrowserIndex(v_iLoop)).Width + frm_Main.sst_General.Width + 45
            frm_Main.tab_Main.Left = 0
            frm_Main.sst_General.Visible = False
        Next v_iLoop
    End If
End Sub

Public Function IsValueInRecordset(m_Value As String, m_Recordet As Recordset) As Boolean
    While Not m_Recordet.EOF
        If m_Value = m_Recordet.Fields(1).Value Then
            IsValueInRecordset = True
            Exit Function
        End If
        m_Recordet.MoveNext
    Wend
    IsValueInRecordset = False
End Function

Public Sub AnalyzeIncomingData(m_Data As String)
    Dim v_iMsg As Integer
    Dim v_iLoop As Integer
    Dim v_sTemp As String
    
    If Left(m_Data, 1) = "$" Then 'Data is a string
      Select Case Mid(m_Data, 2, 3)
        Case "ID:":
            v_iMsg = MsgBox("Connection request from " & Right(m_Data, Len(m_Data) - 4) & ". Let him/her to connects?", vbYesNo + vbInformation, "Connection Request")
            If v_iMsg = vbNo Then
                frm_Main.sok_Server.SendData "%ConnectionRejected"
            Else
                v_sClientID = Right(m_Data, Len(m_Data) - 4)
                frm_Main.sok_Server.SendData "%Connected"
                frm_Main.lbl_Label(5).Caption = "Connected To : " & v_sClientID
            End If
        Case "CT:":
            If frm_Main.tbx_SMsgList.Text <> "" Then
                frm_Main.tbx_SMsgList.Text = frm_Main.tbx_SMsgList.Text & Chr(13) & v_sClientID & ">" & Right(m_Data, Len(m_Data) - 4)
            Else
                frm_Main.tbx_SMsgList.Text = v_sClientID & ">" & Right(m_Data, Len(m_Data) - 4)
            End If
        Case "ST:":
            If frm_Main.tbx_CMsgList.Text <> "" Then
                frm_Main.tbx_CMsgList.Text = frm_Main.tbx_CMsgList.Text & Chr(13) & v_sServerID & ">" & Right(m_Data, Len(m_Data) - 4)
            Else
                frm_Main.tbx_CMsgList.Text = v_sServerID & ">" & Right(m_Data, Len(m_Data) - 4)
            End If
        Case "IP:":
            v_sRemoteIP = Right(m_Data, Len(m_Data) - 4)
            frm_Main.lbl_Label(6).Caption = "Client IP : " & Right(m_Data, Len(m_Data) - 4)
            frm_Main.sok_Server.SendData "#ServerIsConnecting"
        Case "DRV":
            Call AnalyzeDriveString(Right(m_Data, Len(m_Data) - 4))
        Case "FLD":
            Call AnalyzeFolderString(Right(m_Data, Len(m_Data) - 4))
        Case "FIL":
            Call AnalyzeFileString(Right(m_Data, Len(m_Data) - 4))
        Case "FSZ":
            v_lDownloadFileLen = CLng(Right(m_Data, Len(m_Data) - 4))
            frm_FileManager.pbr_Progress.Max = v_lDownloadFileLen
            frm_FileManager.pbr_Progress.Value = 0
            frm_FileManager.lbl_Label(1).Caption = "Downloading..."
            frm_Main.sok_Server.SendData "#StartSendingFile"
            v_bDownloadRequested = True
        End Select
    ElseIf Left(m_Data, 1) = "%" Then 'Data is a returned value
        Select Case Right(m_Data, Len(m_Data) - 1)
        Case "ConnectionRejected":
            MsgBox "Server rejected the connection."
        Case "Connected":
            frm_Main.sok_Client.SendData "$IP:" & frm_Main.sok_Client.LocalIP
            MsgBox "Connected.", vbInformation
            frm_Main.fra_Frame(1).Enabled = True
            frm_Main.fra_Frame(3).Enabled = True
        Case "Disconnecting":
            frm_Main.lbl_Label(5).Caption = "Connected To : No Connection"
            frm_Main.lbl_Label(6).Caption = "Client IP : Not Connected"
            frm_Main.fra_Frame(3).Enabled = False
            frm_Main.fra_Frame(1).Enabled = False
            frm_Main.btn_Connect.Caption = "Connect"
            frm_Main.sok_Client.Close
            frm_Main.sok_Server.Close
            frm_Main.sok_Server.Listen
        Case "ServerCanConnect":
            frm_Main.sok_Client.RemoteHost = v_sRemoteIP
            frm_Main.sok_Client.RemotePort = v_iPortNumber
            frm_Main.sok_Client.Connect
        Case "ServerConnected":
            frm_Main.fra_Frame(4).Enabled = True
            v_sClientID = "Server"
        End Select
    ElseIf Left(m_Data, 1) = "#" Then 'Data is a command
        Select Case Right(m_Data, Len(m_Data) - 1)
        Case "Introduce":
            frm_Main.sok_Client.SendData "$ID:" & frm_Main.tbx_ID.Text
        Case "ServerIsConnecting":
            v_bAutomaticClientConnect = True
            frm_Main.sok_Client.SendData "%ServerCanConnect"
        Case "SendAllDrives":
            v_sTemp = ""
            For v_iLoop = 0 To frm_Main.drv_Drive.ListCount - 1
                v_sTemp = v_sTemp & frm_Main.drv_Drive.List(v_iLoop) & "|"
            Next v_iLoop
            'v_sTemp = Left(v_sTemp, Len(v_sTemp) - 1)
            frm_Main.sok_Client.SendData "$DRV" & v_sTemp
        Case "StartSendingFile":
            v_lDownloadedBytes = 0
            v_sDownloadedData = ""
            v_sTotalDownloadedData = ""
            Open v_sDownloadFileName For Binary As #1
                v_sTemp = String(FileLen(v_sDownloadFileName), " ")
                Get #1, , v_sTemp
            Close #1
            frm_Main.sok_Client.SendData v_sTemp
            DoEvents
        End Select
        
        If Mid(m_Data, 2, 10) = "NavigateTo" Then
            frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).Navigate Right(m_Data, Len(m_Data) - 11)
        End If
    
        If Mid(m_Data, 2, 10) = "MessageBox" Then
            MsgBox Right(m_Data, Len(m_Data) - 11), vbInformation, "Message sent from " & v_sClientID
        End If
        
        If Mid(m_Data, 2, 11) = "ChangeDrive" Then
            frm_Main.dir_Folder.Path = frm_Main.drv_Drive.List(CInt(Right(m_Data, Len(m_Data) - 12)))
            
            v_sTemp = ""
            For v_iLoop = 0 To frm_Main.dir_Folder.ListCount - 1
                v_sTemp = v_sTemp & frm_Main.dir_Folder.List(v_iLoop) & "|"
            Next v_iLoop
            frm_Main.sok_Client.SendData "$FLD" & v_sTemp
        End If
    
        If Mid(m_Data, 2, 12) = "ChangeFolder" Then
            frm_Main.dir_Folder.Path = frm_Main.dir_Folder.List(CInt(Right(m_Data, Len(m_Data) - 13)))
            
            v_sTemp = ""
            For v_iLoop = 0 To frm_Main.dir_Folder.ListCount - 1
                v_sTemp = v_sTemp & frm_Main.dir_Folder.List(v_iLoop) & "|"
            Next v_iLoop
            
            For v_iLoop = 0 To frm_Main.fil_File.ListCount - 1
                v_sTemp = v_sTemp & frm_Main.fil_File.List(v_iLoop) & "~"
            Next v_iLoop
            frm_Main.sok_Client.SendData "$FIL" & v_sTemp
        End If
    
        If Mid(m_Data, 2, 10) = "ChangePath" Then
            frm_Main.dir_Folder.Path = Right(m_Data, Len(m_Data) - 11)
            
            v_sTemp = ""
            For v_iLoop = 0 To frm_Main.dir_Folder.ListCount - 1
                v_sTemp = v_sTemp & Right(frm_Main.dir_Folder.List(v_iLoop), Len(frm_Main.dir_Folder.List(v_iLoop)) - InStrRev(frm_Main.dir_Folder.List(v_iLoop), "\")) & "|"
            Next v_iLoop
            
            For v_iLoop = 0 To frm_Main.fil_File.ListCount - 1
                v_sTemp = v_sTemp & frm_Main.fil_File.List(v_iLoop) & "~"
            Next v_iLoop
            frm_Main.sok_Client.SendData "$FIL" & v_sTemp
        End If
    
        If Mid(m_Data, 2, 8) = "Download" Then
            frm_Main.sok_Client.SendData "$FSZ" & FileLen(Right(m_Data, Len(m_Data) - 9))
            v_sDownloadFileName = Right(m_Data, Len(m_Data) - 9)
        End If
    End If
End Sub

Public Sub AnalyzeDriveString(m_String As String)
    frm_FileManager.cbo_Drive.ComboItems.Clear
    While InStr(m_String, "|") > 0
        frm_FileManager.cbo_Drive.ComboItems.Add , , Left(m_String, InStr(m_String, "|") - 1), 1
        m_String = Right(m_String, Len(m_String) - InStr(m_String, "|"))
    Wend
    frm_FileManager.Show
End Sub

Public Sub AnalyzeFolderString(m_String As String)
    frm_FileManager.lbx_Folder.ListItems.Clear
    frm_FileManager.lbx_Folder.ListItems.Add , , "..", , 4
    While InStr(m_String, "|") > 0
        frm_FileManager.lbx_Folder.ListItems.Add , , Left(m_String, InStr(m_String, "|") - 1), , 4
        m_String = Right(m_String, Len(m_String) - InStr(m_String, "|"))
    Wend
End Sub

Public Sub AnalyzeFileString(m_String As String)
    frm_FileManager.lbx_Folder.ListItems.Clear
    frm_FileManager.lbx_Folder.ListItems.Add , , "..", , 4
    While InStr(m_String, "|") > 0
        frm_FileManager.lbx_Folder.ListItems.Add , , Left(m_String, InStr(m_String, "|") - 1), , 4
        m_String = Right(m_String, Len(m_String) - InStr(m_String, "|"))
    Wend

    frm_FileManager.lbx_File.ListItems.Clear
    While InStr(m_String, "~") > 0
        frm_FileManager.lbx_File.ListItems.Add , , Left(m_String, InStr(m_String, "~") - 1), , 5
        m_String = Right(m_String, Len(m_String) - InStr(m_String, "~"))
    Wend
End Sub

