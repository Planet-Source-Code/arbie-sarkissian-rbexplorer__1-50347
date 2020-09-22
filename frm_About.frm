VERSION 5.00
Begin VB.Form frm_About 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pic_About 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   503
      Picture         =   "frm_About.frx":0000
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   1
      Top             =   840
      Width           =   720
   End
   Begin VB.Label lbl_About 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "RbExplorer version 1.0 copyright(c) 2003 by Arbie Sarkissian"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1463
      TabIndex        =   0
      Top             =   660
      Width           =   3255
   End
End
Attribute VB_Name = "frm_About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
