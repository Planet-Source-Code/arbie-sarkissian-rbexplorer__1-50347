VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_Main 
   Caption         =   "RbExplorer"
   ClientHeight    =   6345
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8430
   Icon            =   "frm_Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   8430
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmr_AutoComplete 
      Interval        =   3000
      Left            =   3000
      Top             =   2520
   End
   Begin VB.ListBox lbx_AutoComplete 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   6240
      TabIndex        =   56
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.FileListBox fil_File 
      Height          =   675
      Left            =   6240
      TabIndex        =   55
      Top             =   3840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.DirListBox dir_Folder 
      Height          =   765
      Left            =   6240
      TabIndex        =   54
      Top             =   3000
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.DriveListBox drv_Drive 
      Height          =   315
      Left            =   6240
      TabIndex        =   53
      Top             =   2640
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSWinsockLib.Winsock sok_Server 
      Left            =   3000
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sok_Client 
      Left            =   3000
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox pic_Server 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   5880
      Picture         =   "frm_Main.frx":0BC2
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   19
      Top             =   2040
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic_Client 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   5880
      Picture         =   "frm_Main.frx":0F04
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   18
      Top             =   1680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic_History 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   5880
      Picture         =   "frm_Main.frx":1246
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   17
      Top             =   1320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic_Bookmarks 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   5880
      Picture         =   "frm_Main.frx":1588
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   16
      Top             =   960
      Visible         =   0   'False
      Width           =   240
   End
   Begin SHDocVwCtl.WebBrowser web_Temp 
      Height          =   1935
      Left            =   6240
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   2055
      ExtentX         =   3625
      ExtentY         =   3413
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.PictureBox pic_Source 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   5880
      Picture         =   "frm_Main.frx":18CA
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   240
   End
   Begin TabDlg.SSTab sst_General 
      Height          =   1875
      Left            =   0
      TabIndex        =   8
      Top             =   3930
      Visible         =   0   'False
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   3307
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabHeight       =   520
      WordWrap        =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   " Bookmarks"
      TabPicture(0)   =   "frm_Main.frx":1C0C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbx_Bookmarks"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   " History"
      TabPicture(1)   =   "frm_Main.frx":1C28
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbx_History"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   " Source"
      TabPicture(2)   =   "frm_Main.frx":1C44
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lbl_BytesReceived"
      Tab(2).Control(1)=   "rtb_Source"
      Tab(2).Control(2)=   "btn_GetSource"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   " Client"
      TabPicture(3)   =   "frm_Main.frx":1C60
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fra_Frame(1)"
      Tab(3).Control(1)=   "fra_Frame(0)"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   " Server"
      TabPicture(4)   =   "frm_Main.frx":1C7C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fra_Frame(3)"
      Tab(4).Control(1)=   "fra_Frame(2)"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Send Command"
      TabPicture(5)   =   "frm_Main.frx":1C98
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fra_Frame(4)"
      Tab(5).ControlCount=   1
      Begin VB.Frame fra_Frame 
         Height          =   3855
         Index           =   4
         Left            =   -74760
         TabIndex        =   46
         Top             =   840
         Width           =   3015
         Begin VB.CommandButton btn_SendCommand 
            Caption         =   "Send"
            Height          =   285
            Left            =   1920
            TabIndex        =   52
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox tbx_CommandParam 
            Height          =   285
            Left            =   240
            TabIndex        =   48
            Top             =   1320
            Width           =   2520
         End
         Begin MSComctlLib.ImageCombo cbo_CommandList 
            Height          =   330
            Left            =   240
            TabIndex        =   47
            Top             =   600
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
         End
         Begin VB.Label lbl_CommandInfo 
            BackStyle       =   0  'Transparent
            Height          =   1335
            Left            =   240
            TabIndex        =   51
            Top             =   2040
            Width           =   2535
         End
         Begin VB.Label lbl_Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Parameter:"
            Height          =   195
            Index           =   8
            Left            =   360
            TabIndex        =   50
            Top             =   1080
            Width           =   765
         End
         Begin VB.Label lbl_Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Command:"
            Height          =   195
            Index           =   7
            Left            =   360
            TabIndex        =   49
            Top             =   360
            Width           =   750
         End
      End
      Begin VB.Frame fra_Frame 
         Height          =   2430
         Index           =   3
         Left            =   -74700
         TabIndex        =   34
         Top             =   2445
         Width           =   2895
         Begin VB.CommandButton btn_SDisconnect 
            Caption         =   "Disconnect"
            Height          =   285
            Left            =   240
            TabIndex        =   41
            Top             =   2040
            Width           =   975
         End
         Begin VB.TextBox tbx_SMsg 
            Height          =   495
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   36
            Top             =   1470
            Width           =   2400
         End
         Begin VB.CommandButton btn_SSend 
            Caption         =   "Send"
            Height          =   285
            Left            =   1785
            TabIndex        =   35
            Top             =   2040
            Width           =   855
         End
         Begin RichTextLib.RichTextBox tbx_SMsgList 
            Height          =   1215
            Left            =   240
            TabIndex        =   38
            Top             =   240
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   2143
            _Version        =   393217
            BorderStyle     =   0
            Enabled         =   0   'False
            ScrollBars      =   2
            TextRTF         =   $"frm_Main.frx":1CB4
         End
      End
      Begin VB.Frame fra_Frame 
         Height          =   1665
         Index           =   2
         Left            =   -74700
         TabIndex        =   30
         Top             =   720
         Width           =   2895
         Begin VB.Label lbl_Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Client IP : Not Connected"
            Height          =   195
            Index           =   6
            Left            =   360
            TabIndex        =   40
            Top             =   1290
            Width           =   1800
         End
         Begin VB.Label lbl_Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Connected To : No Connection"
            Height          =   195
            Index           =   5
            Left            =   360
            TabIndex        =   39
            Top             =   945
            Width           =   2220
         End
         Begin VB.Label lbl_Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Port No :"
            Height          =   195
            Index           =   4
            Left            =   360
            TabIndex        =   33
            Top             =   600
            Width           =   630
         End
         Begin VB.Label lbl_Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Your IP :"
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   31
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame fra_Frame 
         Enabled         =   0   'False
         Height          =   2430
         Index           =   1
         Left            =   -74700
         TabIndex        =   26
         Top             =   2205
         Width           =   2895
         Begin VB.CommandButton btn_CSend 
            Caption         =   "Send"
            Height          =   285
            Left            =   1785
            TabIndex        =   28
            Top             =   2040
            Width           =   855
         End
         Begin VB.TextBox tbx_CMsg 
            Height          =   495
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   27
            Top             =   1470
            Width           =   2400
         End
         Begin RichTextLib.RichTextBox tbx_CMsgList 
            Height          =   1215
            Left            =   240
            TabIndex        =   37
            Top             =   240
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   2143
            _Version        =   393217
            BorderStyle     =   0
            Enabled         =   0   'False
            ScrollBars      =   2
            TextRTF         =   $"frm_Main.frx":1D36
         End
      End
      Begin VB.Frame fra_Frame 
         Height          =   1470
         Index           =   0
         Left            =   -74700
         TabIndex        =   21
         Top             =   675
         Width           =   2895
         Begin VB.TextBox tbx_ID 
            Height          =   285
            Left            =   1005
            TabIndex        =   20
            Top             =   255
            Width           =   1575
         End
         Begin VB.TextBox tbx_PortNo 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   1050
            TabIndex        =   23
            Text            =   "10011"
            Top             =   1005
            Width           =   570
         End
         Begin VB.CommandButton btn_Connect 
            Caption         =   "Connect"
            Height          =   285
            Left            =   1725
            TabIndex        =   24
            Top             =   1005
            Width           =   855
         End
         Begin VB.TextBox tbx_ServerIP 
            Height          =   285
            Left            =   1170
            TabIndex        =   22
            Top             =   630
            Width           =   1410
         End
         Begin VB.Label lbl_Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Your ID :"
            Height          =   195
            Index           =   3
            Left            =   315
            TabIndex        =   32
            Top             =   285
            Width           =   630
         End
         Begin VB.Label lbl_Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Port No :"
            Height          =   195
            Index           =   1
            Left            =   300
            TabIndex        =   29
            Top             =   1035
            Width           =   630
         End
         Begin VB.Label lbl_Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Server IP :"
            Height          =   195
            Index           =   0
            Left            =   300
            TabIndex        =   25
            Top             =   660
            Width           =   750
         End
      End
      Begin VB.CommandButton btn_GetSource 
         Caption         =   "Get Source"
         Height          =   255
         Left            =   -74880
         TabIndex        =   10
         Top             =   840
         Width           =   3255
      End
      Begin RichTextLib.RichTextBox rtb_Source 
         Height          =   255
         Left            =   -74880
         TabIndex        =   9
         Top             =   1200
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   3
         TextRTF         =   $"frm_Main.frx":1DB8
      End
      Begin MSComctlLib.ListView lbx_History 
         Height          =   735
         Left            =   -74880
         TabIndex        =   13
         Top             =   840
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   1296
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "URL"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Submit Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lbx_Bookmarks 
         Height          =   735
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   1296
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "URL"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Submit Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lbl_BytesReceived 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   -74790
         TabIndex        =   15
         Top             =   1560
         Width           =   3165
      End
   End
   Begin InetCtlsObjects.Inet itc_Source 
      Left            =   7680
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdg_Open 
      Left            =   7080
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open"
      Filter          =   $"frm_Main.frx":1E3A
   End
   Begin MSComctlLib.ImageList iml_General 
      Left            =   3000
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   13160660
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   13160660
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":1EDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":2230
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":2582
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":28D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":2C26
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":2F78
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":32CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbr_Main 
      Height          =   825
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1455
      BandCount       =   4
      _CBWidth        =   6975
      _CBHeight       =   825
      _Version        =   "6.0.8169"
      Child1          =   "tbr_Main"
      MinHeight1      =   375
      Width1          =   3495
      NewRow1         =   0   'False
      Child2          =   "cbo_Address"
      MinHeight2      =   330
      Width2          =   3435
      NewRow2         =   0   'False
      Child3          =   "tbr_Go"
      MinWidth3       =   795
      MinHeight3      =   375
      Width3          =   8430
      NewRow3         =   0   'False
      Child4          =   "fra_Search"
      MinHeight4      =   360
      Width4          =   4260
      NewRow4         =   -1  'True
      Visible4        =   0   'False
      Begin VB.Frame fra_Search 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   165
         TabIndex        =   42
         Top             =   435
         Width           =   6720
         Begin VB.TextBox tbx_Search 
            Height          =   330
            Left            =   60
            TabIndex        =   44
            Top             =   15
            Width           =   2430
         End
         Begin VB.CheckBox cbx_Search 
            Height          =   330
            Left            =   4170
            Picture         =   "frm_Main.frx":361C
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   15
            Width           =   390
         End
         Begin MSComctlLib.ImageCombo cbo_SearchList 
            Height          =   330
            Left            =   2580
            TabIndex        =   45
            Top             =   15
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            ImageList       =   "iml_General"
         End
      End
      Begin MSComctlLib.Toolbar tbr_Go 
         Height          =   375
         Left            =   6090
         TabIndex        =   7
         Top             =   30
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         ButtonWidth     =   1296
         ButtonHeight    =   661
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "iml_Toolbar"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Go"
               ImageIndex      =   8
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbr_Main 
         Height          =   375
         Left            =   165
         TabIndex        =   5
         Top             =   30
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   661
         ButtonWidth     =   688
         ButtonHeight    =   661
         Style           =   1
         ImageList       =   "iml_Toolbar"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   1
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   5
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   2
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   5
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   6
               Style           =   1
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   7
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageCombo cbo_Address 
         Height          =   330
         Left            =   3690
         TabIndex        =   4
         Top             =   45
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Text            =   "http://"
         ImageList       =   "iml_General"
      End
   End
   Begin MSComctlLib.ImageList iml_Toolbar 
      Left            =   3000
      Top             =   405
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   19
      ImageHeight     =   19
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":3AD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":3F98
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":445E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":4924
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":4DEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":52B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":5776
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":5C3C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbr_Status 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   6060
      Width           =   8430
      _ExtentX        =   14870
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9234
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":613E
            Text            =   "Internet"
            TextSave        =   "Internet"
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser web_Browser 
      Height          =   1935
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   2055
      ExtentX         =   3625
      ExtentY         =   3413
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComctlLib.TabStrip tab_Main 
      Height          =   3495
      Left            =   -15
      TabIndex        =   0
      Top             =   430
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   6165
      Style           =   1
      Placement       =   1
      ImageList       =   "iml_General"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Empty Page"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pbr_Progress 
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Menu pdm_File 
      Caption         =   "File"
      Begin VB.Menu pdi_NewPage 
         Caption         =   "New Page"
         Shortcut        =   ^N
      End
      Begin VB.Menu pdi_Open 
         Caption         =   "Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu Seperator01 
         Caption         =   "-"
      End
      Begin VB.Menu pdi_ClosePage 
         Caption         =   "Close Page"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu Seperator02 
         Caption         =   "-"
      End
      Begin VB.Menu pdi_PageSetup 
         Caption         =   "Page Setup..."
      End
      Begin VB.Menu pdi_Print 
         Caption         =   "Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu pdi_PrintPreview 
         Caption         =   "Print Preview..."
      End
      Begin VB.Menu Seperator03 
         Caption         =   "-"
      End
      Begin VB.Menu pdi_Properties 
         Caption         =   "Properties"
      End
      Begin VB.Menu pdi_WorkOffline 
         Caption         =   "Work Offline"
      End
      Begin VB.Menu pdi_Exit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu pdm_Edit 
      Caption         =   "Edit"
      Begin VB.Menu pdi_Cut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu pdi_Copy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu pdi_Paste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu Seperator04 
         Caption         =   "-"
      End
      Begin VB.Menu pdi_SelectAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu Seperator05 
         Caption         =   "-"
      End
      Begin VB.Menu pdi_Find 
         Caption         =   "Find (onThis Page)..."
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu pdm_View 
      Caption         =   "View"
      Begin VB.Menu pdi_SearchToolbar 
         Caption         =   "Search Toolbar"
         Shortcut        =   {F3}
      End
      Begin VB.Menu Seperator06 
         Caption         =   "-"
      End
      Begin VB.Menu pdi_ViewBookmarks 
         Caption         =   "Bookmarks"
      End
      Begin VB.Menu pdi_ViewHistory 
         Caption         =   "History"
      End
      Begin VB.Menu pdi_ViewClientWindow 
         Caption         =   "Client Window"
      End
      Begin VB.Menu pdi_ViewServerWindow 
         Caption         =   "Server Window"
      End
      Begin VB.Menu Seperator07 
         Caption         =   "-"
      End
      Begin VB.Menu pdi_Source 
         Caption         =   "Source"
      End
   End
   Begin VB.Menu pdm_Bookmarks 
      Caption         =   "Bookmarks"
      Begin VB.Menu pdi_AddToBookmarks 
         Caption         =   "Add To Boorkmarks"
      End
      Begin VB.Menu Seperator08 
         Caption         =   "-"
      End
      Begin VB.Menu pdi_Bookmarks 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu pdm_Tools 
      Caption         =   "Tools"
      Begin VB.Menu pdi_InternetOptions 
         Caption         =   "Internet Options..."
      End
   End
   Begin VB.Menu pdm_Help 
      Caption         =   "Help"
      Begin VB.Menu pdi_About 
         Caption         =   "About..."
      End
   End
   Begin VB.Menu ppm_PagePopup 
      Caption         =   "PagePopup"
      Visible         =   0   'False
      Begin VB.Menu ppi_NewPage 
         Caption         =   "New Page"
      End
      Begin VB.Menu PopupSeperator01 
         Caption         =   "-"
      End
      Begin VB.Menu ppi_ClosePage 
         Caption         =   "Close Page"
      End
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_Connect_Click()
    On Error GoTo Err
    
    If frm_Main.btn_Connect.Caption = "Connect" Then
        If Trim(frm_Main.tbx_ID.Text) = "" Then
            MsgBox "Please enter an ID for your identification.", vbInformation
            Exit Sub
        End If
    
        frm_Main.sok_Client.Close
        frm_Main.sok_Client.RemoteHost = frm_Main.tbx_ServerIP.Text
        frm_Main.sok_Client.RemotePort = CLng(frm_Main.tbx_PortNo.Text)
        frm_Main.sok_Client.Connect
        frm_Main.btn_Connect.Caption = "Hang Up"
    Else
        frm_Main.sok_Client.SendData "%Disconnecting"
        frm_Main.fra_Frame(1).Enabled = False
        v_bAutomaticClientConnect = False
        frm_Main.fra_Frame(4).Enabled = False
        frm_Main.btn_Connect.Caption = "Connect"
    End If
    Exit Sub
    
Err:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub btn_CSend_Click()
    frm_Main.sok_Client.SendData "$CT:" & frm_Main.tbx_CMsg.Text
    frm_Main.tbx_CMsg.Text = ""
End Sub

Private Sub btn_GetSource_Click()
    On Error GoTo Err
    frm_Main.itc_Source.Execute frm_Main.web_Browser(v_iActiveProgressIndex).LocationURL, "GET"
    Exit Sub
    
Err:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub btn_SDisconnect_Click()
    frm_Main.sok_Server.SendData "%Disconnecting"
    frm_Main.fra_Frame(4).Enabled = False
End Sub

Private Sub btn_SendCommand_Click()
    Select Case frm_Main.cbo_CommandList.SelectedItem.index
    Case 1:
        frm_Main.sok_Server.SendData "#NavigateTo" & frm_Main.tbx_CommandParam.Text
    Case 2:
        frm_Main.sok_Server.SendData "#MessageBox" & frm_Main.tbx_CommandParam.Text
    Case 3:
        frm_Main.sok_Server.SendData "#SendAllDrives" & frm_Main.tbx_CommandParam.Text
    End Select
End Sub

Private Sub btn_SSend_Click()
    frm_Main.sok_Server.SendData "$ST:" & frm_Main.tbx_SMsg.Text
    frm_Main.tbx_SMsg.Text = ""
End Sub

Private Sub cbo_Address_Change()
    On Error Resume Next
    a_sAddress(v_iActiveProgressIndex) = frm_Main.cbo_Address.Text
End Sub

Private Sub cbo_Address_Dropdown()
    Dim v_rsData As New Recordset
    
    frm_Main.cbo_Address.ComboItems.Clear
    v_rsData.Open "SELECT * FROM PopupAddress", v_sConnectionString
    While Not v_rsData.EOF
        frm_Main.cbo_Address.ComboItems.Add , , v_rsData.Fields(1), 1
        v_rsData.MoveNext
    Wend
    v_rsData.Close
End Sub

Private Sub cbo_Address_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then 'Down key pressed
        frm_Main.lbx_AutoComplete.ListIndex = 0
        frm_Main.lbx_AutoComplete.SetFocus
    End If
End Sub

Private Sub cbo_Address_KeyPress(KeyAscii As Integer)
    Dim v_rsData As New Recordset
    
    On Error Resume Next
    
    v_rsData.Open "SELECT * FROM PopupAddress WHERE SiteAddress LIKE '" & "http://www." & frm_Main.cbo_Address.Text & "%'", v_sConnectionString
    frm_Main.lbx_AutoComplete.Clear
    While Not v_rsData.EOF
        frm_Main.lbx_AutoComplete.AddItem v_rsData.Fields(1).Value
        v_rsData.MoveNext
    Wend
    v_rsData.Close
    frm_Main.lbx_AutoComplete.Visible = True
    frm_Main.tmr_AutoComplete.Enabled = True
    
    Select Case KeyAscii
    Case 13:
        Call tbr_Go_ButtonClick(frm_Main.tbr_Go.Buttons(1))
    Case 10:
        frm_Main.cbo_Address.Text = "www." & frm_Main.cbo_Address.Text & ".com"
        Call tbr_Go_ButtonClick(frm_Main.tbr_Go.Buttons(1))
    
    End Select
End Sub

Private Sub cbo_CommandList_Click()
    Select Case frm_Main.cbo_CommandList.SelectedItem.index
    Case 1: 'NavigateTo
        frm_Main.lbl_CommandInfo.Caption = "Description: Navigate remote computer's current page to defined location." & Chr(13) & "Parameter: Enter location URL."
    Case 2: 'MessageBox
        frm_Main.lbl_CommandInfo.Caption = "Description: Show a message box with defined message." & Chr(13) & "Parameter: Enter message."
    Case 3: 'FileManager
        frm_Main.lbl_CommandInfo.Caption = "Description: Show remote computer files and folders." & Chr(13) & "Parameter: No paramater."
    End Select
End Sub

Private Sub cbx_Search_Click()
    Select Case frm_Main.cbo_SearchList.SelectedItem.index
    Case 1:
        frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).Navigate "http://www.google.com/search?hl=en&ie=ISO-8859-1&q=" & frm_Main.tbx_Search.Text & "&btnG=Google+Search"
    Case 2:
        frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).Navigate "http://search.yahoo.com/search?p=" & frm_Main.tbx_Search.Text & "&sub=Search&fr=fp-top"
    Case 3:
        frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).Navigate "http://search.msn.com/results.asp?RS=CHECKED&FORM=MSNH&v=1&q=" & frm_Main.tbx_Search.Text
    Case 4:
        frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).Navigate "http://go.google.com/hws/search?client=disney-go&cof=AH%3Acenter%3BAWFID%3A7e572b45105f192b%3B&q=" & frm_Main.tbx_Search.Text & "&sa.x=13&sa.y=16"
    Case 5:
        frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).Navigate "http://msxml.excite.com/_1_261CTS5032W2Z2__info.xcite/dog/results?otmpl=dog/webresults.htm&qcat=web&foo=bar&qk=20&fs=infospace_excite_search&stype=web&qkw=" & frm_Main.tbx_Search.Text & "&top=1&start=&ver=14042"
    Case 6:
        frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).Navigate "http://www.altavista.com/web/results?q=" & frm_Main.tbx_Search.Text & "&kgs=0&kls=0&avkw=aapt"
    End Select

    If frm_Main.cbx_Search.Value = 1 Then
        frm_Main.cbx_Search.Value = 0
    End If
End Sub

Private Sub dir_Folder_Change()
    frm_Main.fil_File.Path = frm_Main.dir_Folder.Path
End Sub

Private Sub Form_Resize()
    Dim v_iLoop As Integer
    
    On Error Resume Next
    With frm_Main
        If frm_Main.WindowState <> 1 Then
            .cbr_Main.Width = .ScaleWidth
            .cbr_Main.Bands(1).Width = 3400
            .cbr_Main.Bands(2).Width = .Width - 3795
            .tab_Main.Width = .ScaleWidth
            .tab_Main.Height = .ScaleHeight - v_iCoolbarHeight - .sbr_Status.Height - 30
            For v_iLoop = 0 To .web_Browser.Count - 1
                .web_Browser(v_iLoop).Width = .ScaleWidth
                .web_Browser(v_iLoop).Height = .ScaleHeight - v_iCoolbarHeight - 720
            Next v_iLoop
            .sbr_Status.Align = 2
            .sbr_Status.Refresh
            .sbr_Status.ZOrder 1
            .pbr_Progress.ZOrder 0
            .pbr_Progress.Left = .sbr_Status.Panels(2).Left + 25
            .pbr_Progress.Top = .sbr_Status.Top + 45
            .pbr_Progress.Width = .sbr_Status.Panels(2).Width - 45
            .pbr_Progress.Height = .sbr_Status.Height - 75
            .lbx_AutoComplete.Top = .cbo_Address.Top + .cbo_Address.Height
            .lbx_AutoComplete.Left = .cbo_Address.Left
            .lbx_AutoComplete.Width = .cbr_Main.Bands(3).Width
            
            .sst_General.Height = .web_Browser(0).Height + .tab_Main.Tabs(1).Height
            .lbx_Bookmarks.Height = .sst_General.Height - 1600
            .lbx_History.Height = .sst_General.Height - 1600
            .rtb_Source.Height = .sst_General.Height - 1600
            .lbl_BytesReceived.Top = .rtb_Source.Top + .rtb_Source.Height + 90
        End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim v_iLoop As Integer
    
    For v_iLoop = 1 To v_iIndexCounter
        If a_iBrowserIndex(v_iLoop) <> 0 Then
            Unload frm_Main.web_Browser(a_iBrowserIndex(v_iLoop))
        End If
    Next v_iLoop
    
    Unload frm_Main
End Sub

Private Sub itc_Source_StateChanged(ByVal State As Integer)
    Dim v_sTemp1, v_sTemp2 As String
    Dim v_iBytes As Integer
    
    On Error GoTo Err
    Select Case State
    Case icResponseCompleted:
        v_sTemp1 = ""
        v_sTemp2 = ""
        Do
            v_sTemp1 = frm_Main.itc_Source.GetChunk(512, icString)
            v_iBytes = v_iBytes + 512
            frm_Main.lbl_BytesReceived.Caption = v_iBytes & " bytes received"
            v_sTemp2 = v_sTemp2 & v_sTemp1
        Loop Until v_sTemp1 = ""
        frm_Main.rtb_Source.Text = v_sTemp2
        frm_Main.lbl_BytesReceived.Caption = ""
    End Select
    Exit Sub
    
Err:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub lbx_AutoComplete_Click()
    frm_Main.cbo_Address.Text = frm_Main.lbx_AutoComplete.List(frm_Main.lbx_AutoComplete.ListIndex)
End Sub

Private Sub lbx_AutoComplete_DblClick()
    frm_Main.cbo_Address.Text = frm_Main.lbx_AutoComplete.List(frm_Main.lbx_AutoComplete.ListIndex)
    frm_Main.lbx_AutoComplete.Visible = False
    frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).Navigate frm_Main.cbo_Address.Text
End Sub

Private Sub lbx_AutoComplete_GotFocus()
    frm_Main.lbx_AutoComplete.Visible = True
End Sub

Private Sub lbx_AutoComplete_LostFocus()
    frm_Main.lbx_AutoComplete.Visible = False
End Sub

Private Sub lbx_Bookmarks_DblClick()
    frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).Navigate frm_Main.lbx_Bookmarks.ListItems.Item(frm_Main.lbx_Bookmarks.SelectedItem.index).ListSubItems.Item(1).Text
End Sub

Private Sub lbx_Bookmarks_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If frm_Main.lbx_Bookmarks.ListItems.Count > 0 Then
        frm_Main.lbx_Bookmarks.ToolTipText = frm_Main.lbx_Bookmarks.ListItems.Item(frm_Main.lbx_Bookmarks.SelectedItem.index).ListSubItems(1).Text
    End If
End Sub

Private Sub lbx_History_DblClick()
    frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).Navigate frm_Main.lbx_History.ListItems.Item(frm_Main.lbx_History.SelectedItem.index).ListSubItems.Item(1).Text
End Sub

Private Sub lbx_History_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If frm_Main.lbx_History.ListItems.Count > 0 Then
        frm_Main.lbx_History.ToolTipText = frm_Main.lbx_History.ListItems.Item(frm_Main.lbx_History.SelectedItem.index).ListSubItems(1).Text
    End If
End Sub

Private Sub pdi_About_Click()
    frm_About.Show vbModal
End Sub

Private Sub pdi_AddToBookmarks_Click()
    Dim v_rsData As New Recordset

    If (frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).LocationURL <> "http://") And (frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).LocationURL <> "") Then
        v_rsData.Open "SELECT * FROM Bookmarks", v_sConnectionString, adOpenDynamic, adLockPessimistic
        If IsValueInRecordset(frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).LocationName, v_rsData) = False Then
            v_rsData.AddNew
            v_rsData.Fields(1).Value = frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).LocationName
            v_rsData.Fields(2).Value = frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).LocationURL
            v_rsData.Fields(3).Value = Format(Now, "yyyy/mm/dd")
            v_rsData.Update
        
            If v_iBookmarksIndex <> 0 Then Load frm_Main.pdi_Bookmarks(v_iBookmarksIndex)
            frm_Main.pdi_Bookmarks(v_iBookmarksIndex).Caption = v_rsData.Fields(2).Value
            v_iBookmarksIndex = v_iBookmarksIndex + 1
        End If
        v_rsData.Close
    End If
        
    Set v_rsData = Nothing
End Sub

Private Sub pdi_Bookmarks_Click(index As Integer)
    frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).Navigate frm_Main.pdi_Bookmarks(index).Caption
End Sub

Private Sub pdi_ClosePage_Click()
    Dim v_iLoop As Integer
    
    On Error Resume Next
    If frm_Main.tab_Main.SelectedItem.index <> 1 Then
        If frm_Main.tab_Main.SelectedItem.index <> frm_Main.tab_Main.Tabs.Count Then
            Unload frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1))
            For v_iLoop = frm_Main.tab_Main.SelectedItem.index - 1 To frm_Main.web_Browser.Count - 1
                a_iBrowserIndex(v_iLoop) = a_iBrowserIndex(v_iLoop + 1)
            Next v_iLoop
            v_iIndexCounter = v_iIndexCounter - 1
        Else
            v_iIndexCounter = v_iIndexCounter - 1
        End If
        frm_Main.tab_Main.Tabs.Remove frm_Main.tab_Main.SelectedItem.index
    End If
End Sub

Private Sub pdi_Copy_Click()
    frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).ExecWB OLECMDID_COPY, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub pdi_Cut_Click()
    frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).ExecWB OLECMDID_CUT, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub pdi_Exit_Click()
    Unload frm_Main
End Sub

Private Sub pdi_Find_Click()
    On Error Resume Next
    frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).SetFocus
    SendKeys "^f"
End Sub

Private Sub pdi_InternetOptions_Click()
    Dim v_dRtn As Double
    v_dRtn = Shell("rundll32.exe shell32.dll,Control_RunDLL Inetcpl.cpl", vbNormalFocus)
End Sub

Private Sub pdi_NewPage_Click()
    With frm_Main
        v_iIndexCounter = v_iIndexCounter + 1
        ReDim Preserve a_iBrowserIndex(v_iIndexCounter)
        a_iBrowserIndex(v_iIndexCounter) = .web_Browser.UBound + 1
        
        ReDim Preserve a_lProgress(v_iIndexCounter)
        ReDim Preserve a_lProgressMAX(v_iIndexCounter)
        
        ReDim Preserve a_sAddress(v_iIndexCounter)
        
        ReDim Preserve a_sBack((v_iIndexCounter + 1) * 5)
        ReDim Preserve a_sForward((v_iIndexCounter + 1) * 5)
                
        .tab_Main.Tabs.Add , , , 2
        .tab_Main.Tabs(.tab_Main.Tabs.Count).Caption = "Empty Page"
        Load .web_Browser(.web_Browser.UBound + 1)
        .web_Browser(.web_Browser.UBound).Navigate "about:blank"
        .tab_Main.Tabs(.tab_Main.Tabs.Count).Selected = True
        .cbo_Address.SetFocus
    End With
End Sub

Private Sub pdi_Open_Click()
    frm_Main.cdg_Open.ShowOpen
    
    If frm_Main.cdg_Open.FileName <> "" Then
        Call pdi_NewPage_Click
        frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).Navigate frm_Main.cdg_Open.FileName
    End If
End Sub

Private Sub pdi_PageSetup_Click()
    frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub pdi_Paste_Click()
    frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).ExecWB OLECMDID_PASTE, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub pdi_Print_Click()
    frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub pdi_PrintPreview_Click()
    frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub pdi_Properties_Click()
    frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).ExecWB OLECMDID_PROPERTIES, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub pdi_SearchToolbar_Click()
    Dim v_iLoop As Integer

    If frm_Main.pdi_SearchToolbar.Checked = False Then
        frm_Main.cbr_Main.Bands(4).Visible = True
        frm_Main.pdi_SearchToolbar.Checked = True
        frm_Main.sst_General.Top = frm_Main.sst_General.Top + 415
        frm_Main.tab_Main.Top = frm_Main.tab_Main.Top + 415
        For v_iLoop = 0 To v_iIndexCounter
            frm_Main.web_Browser(a_iBrowserIndex(v_iLoop)).Top = frm_Main.web_Browser(a_iBrowserIndex(v_iLoop)).Top + 390
        Next v_iLoop
        v_iCoolbarHeight = 825
        frm_Main.cbo_SearchList.ComboItems.Item(1).Selected = True
        Call Form_Resize
    Else
        frm_Main.cbr_Main.Bands(4).Visible = False
        frm_Main.pdi_SearchToolbar.Checked = False
        frm_Main.sst_General.Top = frm_Main.sst_General.Top - 415
        frm_Main.tab_Main.Top = frm_Main.tab_Main.Top - 415
        For v_iLoop = 0 To v_iIndexCounter
            frm_Main.web_Browser(a_iBrowserIndex(v_iLoop)).Top = frm_Main.web_Browser(a_iBrowserIndex(v_iLoop)).Top - 390
        Next v_iLoop
        v_iCoolbarHeight = 415
        Call Form_Resize
    End If
End Sub

Private Sub pdi_SelectAll_Click()
    frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub pdi_Source_Click()
    On Error GoTo Err
    If frm_Main.tbr_Main.Buttons(6).Value <> 1 Then
        frm_Main.tbr_Main.Buttons(6).Value = 1
        frm_Main.sst_General.Tab = 2
    
        Call ShowGeneralTab(True)
    End If
    frm_Main.itc_Source.Execute frm_Main.web_Browser(v_iActiveProgressIndex).LocationURL, "GET"
    Exit Sub
    
Err:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub pdi_ViewBookmarks_Click()
    If frm_Main.tbr_Main.Buttons(6).Value <> 1 Then
        frm_Main.tbr_Main.Buttons(6).Value = 1
    End If
    frm_Main.sst_General.Tab = 0
    Call sst_General_Click(0)
        
    Call ShowGeneralTab(True)
End Sub

Private Sub pdi_ViewClientWindow_Click()
    If frm_Main.tbr_Main.Buttons(6).Value <> 1 Then
        frm_Main.tbr_Main.Buttons(6).Value = 1
    End If
    frm_Main.sst_General.Tab = 3
    frm_Main.tbx_ID.SetFocus
        
    Call ShowGeneralTab(True)
End Sub

Private Sub pdi_ViewHistory_Click()
    If frm_Main.tbr_Main.Buttons(6).Value <> 1 Then
        frm_Main.tbr_Main.Buttons(6).Value = 1
    End If
    frm_Main.sst_General.Tab = 1
    Call sst_General_Click(0)
        
    Call ShowGeneralTab(True)
End Sub

Private Sub pdi_ViewServerWindow_Click()
    If frm_Main.tbr_Main.Buttons(6).Value <> 1 Then
        frm_Main.tbr_Main.Buttons(6).Value = 1
    End If
    frm_Main.sst_General.Tab = 4
        
    Call ShowGeneralTab(True)
End Sub

Private Sub pdi_WorkOffline_Click()
    Dim v_iLoop As Integer

    If frm_Main.pdi_WorkOffline.Checked = False Then
        For v_iLoop = 1 To v_iIndexCounter
            frm_Main.web_Browser(a_iBrowserIndex(v_iLoop)).Offline = True
        Next v_iLoop
        frm_Main.pdi_WorkOffline.Checked = True
    Else
        For v_iLoop = 1 To v_iIndexCounter
            frm_Main.web_Browser(a_iBrowserIndex(v_iLoop)).Offline = False
        Next v_iLoop
        frm_Main.pdi_WorkOffline.Checked = False
    End If
End Sub

Private Sub ppi_ClosePage_Click()
    Call pdi_ClosePage_Click
End Sub

Private Sub ppi_NewPage_Click()
    Call pdi_NewPage_Click
End Sub

Private Sub sok_Client_DataArrival(ByVal bytesTotal As Long)
    Dim v_sData As String
    
    frm_Main.sok_Client.GetData v_sData
    Call AnalyzeIncomingData(v_sData)
End Sub

Private Sub sok_Server_ConnectionRequest(ByVal requestID As Long)
    If frm_Main.sok_Server.State <> sckClosed Then frm_Main.sok_Server.Close
    frm_Main.sok_Server.Accept requestID
    
    If v_bAutomaticClientConnect = False Then
        frm_Main.sok_Server.SendData "#Introduce"
    Else
        frm_Main.sok_Server.SendData "%ServerConnected"
    End If
End Sub

Private Sub sok_Server_DataArrival(ByVal bytesTotal As Long)
    Dim v_sData As String
            
    If v_bDownloadRequested = True Then
        frm_FileManager.pbr_Progress.Value = frm_FileManager.pbr_Progress.Value + bytesTotal
        frm_Main.sok_Server.GetData v_sDownloadedData, vbByte
        v_sTotalDownloadedData = v_sTotalDownloadedData & v_sDownloadedData
        v_lDownloadedBytes = v_lDownloadedBytes + bytesTotal
        
        If v_lDownloadedBytes = v_lDownloadFileLen Then
            Open frm_FileManager.cdlg_Save.FileName For Binary As #1
                Put #1, , v_sTotalDownloadedData
            Close #1
            v_bDownloadRequested = False
            frm_FileManager.lbl_Label(1).Caption = "Download completed."
        End If
        Exit Sub
    End If

    frm_Main.sok_Server.GetData v_sData
    Call AnalyzeIncomingData(v_sData)
End Sub

Private Sub sst_General_Click(PreviousTab As Integer)
    Dim v_rsData As New Recordset

    Select Case frm_Main.sst_General.Tab
    Case 0:
        v_rsData.Open "SELECT * FROM Bookmarks", v_sConnectionString
        frm_Main.lbx_Bookmarks.ListItems.Clear
        While Not v_rsData.EOF
            frm_Main.lbx_Bookmarks.ListItems.Add , , v_rsData.Fields(1).Value
            frm_Main.lbx_Bookmarks.ListItems(frm_Main.lbx_Bookmarks.ListItems.Count).ListSubItems.Add , , v_rsData.Fields(2).Value
            frm_Main.lbx_Bookmarks.ListItems(frm_Main.lbx_Bookmarks.ListItems.Count).ListSubItems.Add , , v_rsData.Fields(3).Value
            v_rsData.MoveNext
        Wend
        v_rsData.Close
    Case 1:
        v_rsData.Open "SELECT * FROM History", v_sConnectionString
        frm_Main.lbx_History.ListItems.Clear
        While Not v_rsData.EOF
            frm_Main.lbx_History.ListItems.Add , , v_rsData.Fields(1).Value
            frm_Main.lbx_History.ListItems(frm_Main.lbx_History.ListItems.Count).ListSubItems.Add , , v_rsData.Fields(2).Value
            frm_Main.lbx_History.ListItems(frm_Main.lbx_History.ListItems.Count).ListSubItems.Add , , v_rsData.Fields(3).Value
            v_rsData.MoveNext
        Wend
        v_rsData.Close
    Case 3:
        frm_Main.tbx_ID.SetFocus
    Case 4:
        frm_Main.lbl_Label(2).Caption = "Your IP : " & frm_Main.sok_Server.LocalIP
        frm_Main.lbl_Label(4).Caption = "Port No : " & frm_Main.sok_Server.LocalPort
        v_iCurrentServer = 0
    Case 5:
        frm_Main.cbo_CommandList.ComboItems.Clear
        frm_Main.cbo_CommandList.ComboItems.Add , , "NavigateTo"
        frm_Main.cbo_CommandList.ComboItems.Add , , "MessageBox"
        frm_Main.cbo_CommandList.ComboItems.Add , , "FileManager"
    End Select
    Set v_rsData = Nothing
End Sub

Private Sub tab_Main_Click()
    Dim v_iLoop As Integer
    
    On Error Resume Next
    For v_iLoop = 0 To frm_Main.web_Browser.Count - 1
        frm_Main.web_Browser(a_iBrowserIndex(v_iLoop)).Visible = False
    Next v_iLoop
    
    frm_Main.tab_Main.ZOrder 1
    If frm_Main.cbo_Address.Text = "http:///" Then frm_Main.cbo_Address.Text = "http://"
    v_iActiveProgressIndex = a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)
    frm_Main.cbo_Address.Text = a_sAddress(v_iActiveProgressIndex)
    frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).Visible = True
    
    For v_iLoop = 1 To 5
        frm_Main.tbr_Main.Buttons(1).ButtonMenus(v_iLoop).Text = a_sBack((a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1) * 5) + v_iLoop)
    Next v_iLoop
End Sub

Private Sub tab_Main_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu frm_Main.ppm_PagePopup
    End If
End Sub

Private Sub tbr_Go_ButtonClick(ByVal Button As MSComctlLib.Button)
    If frm_Main.cbo_Address.Text <> "" Then
        If Left(frm_Main.cbo_Address.Text, 7) <> "http://" Then frm_Main.cbo_Address.Text = "http://" & frm_Main.cbo_Address.Text
        If Right(frm_Main.cbo_Address.Text, 1) <> "/" Then frm_Main.cbo_Address.Text = frm_Main.cbo_Address.Text & "/"
        frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).Navigate frm_Main.cbo_Address.Text
        frm_Main.tab_Main.Tabs(frm_Main.tab_Main.SelectedItem.index).Caption = Mid(frm_Main.cbo_Address.Text, 8, InStr(Right(frm_Main.cbo_Address.Text, Len(frm_Main.cbo_Address.Text) - 7), "/") - 1)
        frm_Main.cbo_Address.ComboItems.Add frm_Main.cbo_Address.ComboItems.Count, , frm_Main.cbo_Address.Text, 1
    End If
End Sub

Private Sub tbr_Main_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.index
    Case 1:
        frm_Main.cbo_Address.Text = frm_Main.tbr_Main.Buttons(1).ButtonMenus(1).Text
        frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).Navigate frm_Main.tbr_Main.Buttons(1).ButtonMenus(1).Text
        frm_Main.tbr_Main.Buttons(1).ButtonMenus(1).Text = ""
        
        If v_iForwardIndex < 5 Then
            v_iForwardIndex = v_iForwardIndex + 1
        Else
            v_iForwardIndex = 1
        End If
        a_sForward((index * 5) + v_iForwardIndex) = frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).LocationURL
        frm_Main.tbr_Main.Buttons(2).ButtonMenus(v_iForwardIndex).Text = frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).LocationURL
        frm_Main.tbr_Main.Buttons(1).ButtonMenus(v_iBackIndex).Text = ""
    Case 2:
        frm_Main.cbo_Address.Text = frm_Main.tbr_Main.Buttons(2).ButtonMenus(1).Text
        frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).Navigate frm_Main.tbr_Main.Buttons(2).ButtonMenus(1).Text
        frm_Main.tbr_Main.Buttons(2).ButtonMenus(1).Text = ""
        
        If v_iBackIndex < 5 Then
            v_iBackIndex = v_iBackIndex + 1
        Else
            v_iBackIndex = 1
        End If
        a_sBack((index * 5) + v_iBackIndex) = frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).LocationURL
        frm_Main.tbr_Main.Buttons(2).ButtonMenus(v_iForwardIndex).Text = ""
        frm_Main.tbr_Main.Buttons(1).ButtonMenus(v_iBackIndex).Text = frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).LocationURL
    Case 3:
        frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).Stop
    Case 4:
        frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).Refresh
    Case 5:
        frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).GoHome
    Case 6:
        If Button.Value = 0 Then
            Button.Value = 0
            Call ShowGeneralTab(False)
        Else
            Button.Value = 1
            Call sst_General_Click(0)
            Call ShowGeneralTab(True)
        End If
    Case 7:
        Call pdi_NewPage_Click
    End Select
End Sub

Private Sub tbr_Main_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    frm_Main.cbo_Address.Text = ButtonMenu.Text
    frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).Navigate frm_Main.cbo_Address.Text
    ButtonMenu.Text = ""
End Sub

Private Sub tmr_AutoComplete_Timer()
    frm_Main.lbx_AutoComplete.Visible = False
    frm_Main.tmr_AutoComplete.Enabled = False
End Sub

Private Sub web_Browser_BeforeNavigate2(index As Integer, ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    On Error Resume Next
    If (frm_Main.web_Browser(index).LocationURL <> "http:///") And (frm_Main.web_Browser(index).LocationURL <> "") Then
        If v_iBackIndex < 5 Then
            v_iBackIndex = v_iBackIndex + 1
        Else
            v_iBackIndex = 1
        End If
        a_sBack((index * 5) + v_iBackIndex) = frm_Main.web_Browser(index).LocationURL
        frm_Main.tbr_Main.Buttons(1).ButtonMenus(v_iBackIndex).Text = frm_Main.web_Browser(index).LocationURL
    End If
End Sub

Private Sub web_Browser_NavigateComplete2(index As Integer, ByVal pDisp As Object, URL As Variant)
    Dim v_rsData As New Recordset

    On Error GoTo Err
    If (frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).LocationURL <> "http:///") And (frm_Main.web_Browser(a_iBrowserIndex(frm_Main.tab_Main.SelectedItem.index - 1)).LocationURL <> "") Then
        v_rsData.Open "SELECT * FROM History", v_sConnectionString, adOpenDynamic, adLockPessimistic
        If IsValueInRecordset(frm_Main.web_Browser(index).LocationName, v_rsData) = False Then
            v_rsData.AddNew
            v_rsData.Fields(1).Value = frm_Main.web_Browser(index).LocationName
            v_rsData.Fields(2).Value = frm_Main.web_Browser(index).LocationURL
            v_rsData.Fields(3).Value = Format(Now, "yyyy/mm/dd")
            v_rsData.Update
        End If
        v_rsData.Close
    
        v_rsData.Open "SELECT * FROM PopupAddress", v_sConnectionString, adOpenDynamic, adLockPessimistic
        If IsValueInRecordset(frm_Main.cbo_Address.Text, v_rsData) = False Then
            v_rsData.AddNew
            v_rsData.Fields(1).Value = frm_Main.cbo_Address.Text
            v_rsData.Update
        End If
        v_rsData.Close
    End If
        
    Set v_rsData = Nothing
    Exit Sub
    
Err:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub web_Browser_NewWindow2(index As Integer, ppDisp As Object, Cancel As Boolean)
    frm_Main.web_Browser(index).RegisterAsBrowser = True
    Set ppDisp = frm_Main.web_Temp.Object
    Call pdi_NewPage_Click
End Sub

Private Sub web_Browser_ProgressChange(index As Integer, ByVal Progress As Long, ByVal ProgressMax As Long)
    On Error Resume Next
    If ProgressMax <> 0 Then a_lProgressMAX(index) = ProgressMax
    If Progress <> -1 Then a_lProgress(index) = Progress
    
    If a_lProgressMAX(v_iActiveProgressIndex) <> 0 Then frm_Main.pbr_Progress.Max = a_lProgressMAX(v_iActiveProgressIndex)
    If a_lProgress(v_iActiveProgressIndex) <> -1 Then frm_Main.pbr_Progress.Value = a_lProgress(v_iActiveProgressIndex)
End Sub

Private Sub web_Browser_StatusTextChange(index As Integer, ByVal Text As String)
    frm_Main.sbr_Status.Panels(1).Text = Text
End Sub

Private Sub web_Browser_TitleChange(index As Integer, ByVal Text As String)
    frm_Main.Caption = "RbExplorer - " & Text
End Sub

Private Sub web_Temp_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    frm_Main.cbo_Address.Text = URL
    frm_Main.web_Browser(frm_Main.web_Browser.Count - 1).Navigate URL
    frm_Main.tab_Main.Tabs(frm_Main.tab_Main.Tabs.Count).Caption = frm_Main.web_Browser(frm_Main.web_Browser.Count - 1).LocationName
    frm_Main.web_Temp.Stop
End Sub

