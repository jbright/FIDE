VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMain 
   Caption         =   "FIDE: Fluke 9010A Integrated Development Environment"
   ClientHeight    =   9150
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11070
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9150
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgPrint 
      Left            =   1080
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picH2Splitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   50
      Left            =   3240
      LinkTimeout     =   65
      ScaleHeight     =   19.595
      ScaleMode       =   0  'User
      ScaleWidth      =   48672
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4200
      Visible         =   0   'False
      Width           =   4680
   End
   Begin RichTextLib.RichTextBox txtActiveFile 
      Height          =   2175
      Left            =   3000
      TabIndex        =   15
      Top             =   1080
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   3836
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      OLEDragMode     =   0
      OLEDropMode     =   1
      TextRTF         =   $"frmMain.frx":173A
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
   Begin MSComctlLib.ImageList ilSmallToolBar 
      Left            =   240
      Top             =   7200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   10
      ImageHeight     =   10
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17BA
            Key             =   "SmallSave"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18CC
            Key             =   "SmallBigger"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19DE
            Key             =   "SmallSmaller"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AF0
            Key             =   "SmallEmpty"
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser wbHelpContext 
      Height          =   1335
      Left            =   7080
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   2415
      ExtentX         =   4260
      ExtentY         =   2355
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
   Begin MSComctlLib.ProgressBar pbSend 
      Height          =   150
      Left            =   0
      TabIndex        =   12
      Top             =   8730
      Visible         =   0   'False
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer ReceiverTimer 
      Interval        =   250
      Left            =   240
      Top             =   6000
   End
   Begin MSCommLib.MSComm SerialConnection 
      Left            =   120
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.TextBox txtSerialOutput 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtOutput 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   3120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   2400
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   624
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   65
   End
   Begin VB.PictureBox picHSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   50
      Left            =   0
      LinkTimeout     =   65
      ScaleHeight     =   19.595
      ScaleMode       =   0  'User
      ScaleWidth      =   48672
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   4680
   End
   Begin MSComctlLib.ImageList ilProjectList 
      Left            =   120
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C02
            Key             =   "FolderOpen"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F90
            Key             =   "FolderClosed"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2311
            Key             =   "Project"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":26A3
            Key             =   "Notes"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A81
            Key             =   "Script"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvProjectList 
      Height          =   1320
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   705
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   2328
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   39
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ilProjectList"
      Appearance      =   1
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   15
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NewFolder"
            Object.ToolTipText     =   "Create a new folder"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EditFolder"
            Object.ToolTipText     =   "Edit the name of this folder"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DeleteFolder"
            Object.ToolTipText     =   "Delete this folder and all its contents"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FileNew"
            Object.ToolTipText     =   "Create a new text or script file"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FileSave"
            Object.ToolTipText     =   "Save current file"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Compile"
            Object.ToolTipText     =   "Compile current test script"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Properties"
            Object.ToolTipText     =   "Set serial port settings"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SerialReceive"
            Object.ToolTipText     =   "Open serial port to receive data from the 9010A"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SerialSend"
            Object.ToolTipText     =   "Send compiled file to the 9010A"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BinarySend"
            Object.ToolTipText     =   "Send a pre-compiled file to the 9010A base unit"
            ImageIndex      =   17
         EndProperty
      EndProperty
      Begin MSComctlLib.Toolbar tbTabToolBar 
         Height          =   330
         Left            =   10560
         TabIndex        =   11
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imlToolbarIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "CloseCurrentFile"
               Object.ToolTipText     =   "Close the current file"
               ImageIndex      =   11
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   8880
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12832
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   240
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1080
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E19
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F2B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34C5
            Key             =   "ReceiveFrom"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A5F
            Key             =   "SendTo"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3FF9
            Key             =   "Settings"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4593
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":46A5
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":47B7
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":48C9
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":49DB
            Key             =   "ClearLog"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F75
            Key             =   "CloseWindow"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":550F
            Key             =   "CompileFile"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5AA9
            Key             =   "DeleteFolder"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6043
            Key             =   "EditFolder"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":65DD
            Key             =   "AddFolder"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6B77
            Key             =   "SaveFile"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7111
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvFiles 
      CausesValidation=   0   'False
      Height          =   1095
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   1931
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      Icons           =   "ilProjectList"
      SmallIcons      =   "ilProjectList"
      ColHdrIcons     =   "ilProjectList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   88194
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.TabStrip tsFiles 
      Height          =   3015
      Left            =   2880
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5318
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tsOutput 
      Height          =   2775
      Left            =   3000
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5280
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4895
      Placement       =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Messages"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Serial Port"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbOutputTabs 
      Height          =   240
      Left            =   9000
      TabIndex        =   14
      Top             =   5040
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   423
      ButtonWidth     =   450
      ButtonHeight    =   423
      Style           =   1
      ImageList       =   "ilSmallToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "SaveOutput"
            Object.ToolTipText     =   "Save this output"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "ClearOuput"
            Object.ToolTipText     =   "Clear the output area"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MaximizeOutput"
            Object.ToolTipText     =   "Maximize this window area"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MinimizeOutput"
            Object.ToolTipText     =   "Minimize this window area"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.Image imgSplitter 
      Height          =   4785
      Left            =   2280
      MousePointer    =   9  'Size W E
      Top             =   720
      Width           =   100
   End
   Begin VB.Image imgHSplitter 
      Height          =   100
      Left            =   0
      MousePointer    =   7  'Size N S
      Top             =   3360
      Width           =   4785
   End
   Begin VB.Image imgH2Splitter 
      Height          =   240
      Left            =   3240
      MousePointer    =   7  'Size N S
      Top             =   3960
      Width           =   4785
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "Save &All"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuFileRename 
         Caption         =   "Rena&me"
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuProject 
      Caption         =   "&Project"
      Begin VB.Menu mnuProjectSetRoot 
         Caption         =   "&Set Root Directory"
      End
      Begin VB.Menu mnuProjectSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProjectAddFolder 
         Caption         =   "Add Folder"
      End
      Begin VB.Menu mnuProjectEditName 
         Caption         =   "Edit Folder Name"
      End
      Begin VB.Menu mnuProjectDeleteFolder 
         Caption         =   "Delete Folder"
      End
   End
   Begin VB.Menu mnuScript 
      Caption         =   "&Script"
      Begin VB.Menu mnuScriptCompile 
         Caption         =   "&Compile"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuScriptCompileSend 
         Caption         =   "Compile && Send"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuScriptBinarySend 
         Caption         =   "Send &Binary File"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuScriptSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScriptOpen 
         Caption         =   "Open Port to Receive"
      End
      Begin VB.Menu mnuScriptDivider2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScriptPromptBefore 
         Caption         =   "Prompt Before Sending"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSepN 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScriptColors 
         Caption         =   "Use Syntax Highlighting"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsETERM 
         Caption         =   "Insert &Listing file Code"
      End
      Begin VB.Menu mnuToolsROMSignature 
         Caption         =   "Insert &ROM Signature Code"
      End
      Begin VB.Menu mnuToolsSepN 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsROMDiff 
         Caption         =   "ROM Set Identifier (from binaries)"
      End
      Begin VB.Menu mnuToolsCalc 
         Caption         =   "Binary to Hex Converter"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&Help"
      Begin VB.Menu mnuSearchForTopic 
         Caption         =   "&Search Help Topic"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuViewDyanmicHelp 
         Caption         =   "&Use Dynamic Help"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnViewSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpView9LC 
         Caption         =   "&View 9LC Commands"
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About FIDE"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long


Private mbMoving As Boolean
Private mbHMoving As Boolean
Private mbH2Moving As Boolean
Public OutputMinimized As Boolean
Private nLastSplitterPos As Single
Private nTreeNodeCount As Integer

Private m_activefilecollection As Collection
Private m_previousActiveFileEditor As FileEditor
Private Const WM_PASTE = &H302
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1
Private Const sglSplitLimit = 500
Private Const sglHSplitLimit = 500
Private Const sglH2SplitLimit = 500

Private UIFileDirty As Boolean
Private UIFileReload As Boolean
Public LockEditorWindow As Boolean

' The follow three are tests
Private hOldWindow As Long
Private hHookedWindow As Long
Public hMsgWindow As Long


    
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Current directory that we are on
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Property Get SelectedDirectory() As String
    If tvProjectList.SelectedItem Is Nothing Then
        ' if we don't have anything selected, then we
        ' will assume the root project.
        SelectedDirectory = frmRootLocation.ProjectPath
    Else
        SelectedDirectory = tvProjectList.SelectedItem.Key
    End If
End Property


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' returns a collection of active files (file editor modules)
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Property Get ActiveFileCollection() As Collection

    If (m_activefilecollection Is Nothing) Then
        Set m_activefilecollection = New Collection
    End If
    Set ActiveFileCollection = m_activefilecollection
    
End Property

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Current file editor that we are working with
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Property Get ActiveFileEditor() As FileEditor
    Dim cActiveTab As MSComctlLib.Tab
    Set cActiveTab = ActiveTab
    If cActiveTab Is Nothing Then
        Set ActiveFileEditor = Nothing
    Else
        Set ActiveFileEditor = ActiveFileCollection(cActiveTab.Key)
    End If
End Property

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' PreviousActiveFileEditor
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Property Set PreviousActiveFileEditor(nNewObject As FileEditor)
    Set m_previousActiveFileEditor = nNewObject
End Property
Public Property Get PreviousActiveFileEditor() As FileEditor
    Set PreviousActiveFileEditor = m_previousActiveFileEditor
End Property


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Current active tab
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Property Get ActiveTab() As MSComctlLib.Tab
    If tsFiles.SelectedItem Is Nothing Then
        Set ActiveTab = Nothing
    Else
        Set ActiveTab = tsFiles.Tabs(tsFiles.SelectedItem.Index)
    End If
End Property



' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: ToggleFileMenu
' PURPOSE: Close, Delete, Save, and Save All all get set at the same time
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub ToggleFileMenu(bNewValue As Boolean)
    mnuFileSave.Enabled = bNewValue
    mnuFilePrint.Enabled = bNewValue
    mnuFileSaveAll.Enabled = bNewValue
    mnuFileClose.Enabled = bNewValue
    mnuFileDelete.Enabled = bNewValue
    mnuFileRename.Enabled = False       ' Not implemented yet
    mnuScriptCompile.Enabled = bNewValue
    mnuEditSelectAll.Enabled = bNewValue
    mnuToolsETERM.Enabled = bNewValue
    mnuToolsROMDiff.Enabled = bNewValue
    mnuToolsROMSignature.Enabled = bNewValue
    
    mnuScriptCompileSend.Enabled = bNewValue And frmOptions.SerialSettingsSet
    
    tbToolBar.Buttons("FileSave").Enabled = bNewValue
    tbToolBar.Buttons("Compile").Enabled = bNewValue
    tbToolBar.Buttons("SerialSend").Enabled = bNewValue And frmOptions.SerialSettingsSet
    
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: Form_Load
' PURPOSE: Built in call. We get this when this form first starts up.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Form_Load()
On Error GoTo ErrLoading

    OutputLog.AddTraceLine "Loading FIDE..."
    
    ' For color sytax highlighting
    ColorSyntax.LoadSyntax App.Path & "\Internal\keywords.txt"
    
    tbTabToolBar.Buttons(1).Enabled = False
    
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    
    imgSplitter.Left = GetSetting(App.Title, "Settings", "VertBar", 2200)
    imgHSplitter.Top = GetSetting(App.Title, "Settings", "HorizBar", 5500)
    imgH2Splitter.Top = GetSetting(App.Title, "Settings", "Horiz2Bar", 7000)
    
    picSplitter.Left = imgSplitter.Left
    picHSplitter.Top = imgHSplitter.Top
    picH2Splitter.Top = imgH2Splitter.Top
    
    OutputMinimized = False
    
    mnuViewDyanmicHelp.Checked = GetSetting(App.Title, "Settings", "DynamicHelpOn", True)
    mnuScriptPromptBefore.Checked = GetSetting(App.Title, "Settings", "PromptBeforeSending", True)
    mnuScriptColors.Checked = GetSetting(App.Title, "Settings", "ScriptColors", True)
    
    ' Load our tree.
    LoadTree frmRootLocation.ProjectPath
    
    tvProjectList.Nodes(1).Selected = True
    tvProjectList_Click

    OutputLog.AddTraceLine "Done!"
    SetTopic "Welcome.htm"
    
    ' Set the default menu items
    ' These items are only valid when at least one file
    ' is opened.
    ToggleFileMenu False
    tbToolBar.Buttons("SerialReceive").Enabled = frmOptions.SerialSettingsSet
    UIFileReload = False
    UIFileDirty = False
    LockEditorWindow = False
    Exit Sub
    
ErrLoading:
    MsgBox "ERROR! Could not properly load the application. " & vbCrLf & vbCrLf & Err.Description
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: Form_QueryUnload
' PURPOSE: Built in call. Only hue can prevent data loss
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim objFileEditor As FileEditor
    Dim answer As VbMsgBoxResult
    Dim i As Integer
    For i = 1 To ActiveFileCollection.Count
        Set objFileEditor = ActiveFileCollection.Item(i)
        If objFileEditor.FileHasChanged Then
        
            answer = _
                MsgBox("File: " & vbCrLf & vbCrLf & _
                    objFileEditor.FileNameAndDirectory & vbCrLf & vbCrLf & _
                    "has changed. Do you wish to save changes?", vbYesNoCancel, "Save Changes?")
                    
            If answer = vbCancel Then
                Cancel = 1
                Exit Sub
            ElseIf answer = vbYes Then
                objFileEditor.OnSave
            End If
        
        End If
        
    Next
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: Form_Unload
' PURPOSE: Called when this is about to exit. We'll want to unload our friends.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer

    ' close the comm port if it's still open
    ClosePort
    
    Set PreviousActiveFileEditor = Nothing
    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    
    SaveSetting App.Title, "Settings", "VertBar", imgSplitter.Left
    SaveSetting App.Title, "Settings", "HorizBar", imgHSplitter.Top
    If Not (OutputMinimized) Then
        SaveSetting App.Title, "Settings", "Horiz2Bar", imgH2Splitter.Top
    End If
    
    SaveSetting App.Title, "Settings", "DynamicHelpOn", mnuViewDyanmicHelp.Checked
    SaveSetting App.Title, "Settings", "PromptBeforeSending", mnuScriptPromptBefore.Checked
    SaveSetting App.Title, "Settings", "ScriptColors", mnuScriptColors.Checked
    
    Dim objFileEditor As FileEditor
    For i = 1 To ActiveFileCollection.Count
        Set objFileEditor = ActiveFileCollection.Item(i)
        objFileEditor.Unload
    Next
    
End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: Form_Resize
' PURPOSE: This window is being resized. We need to scale properly
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < 3000 Then Me.Width = 3000
    If Me.Height < 3000 Then Me.Height = 3000
    SizeControls
End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: LoadTree
' PURPOSE: Load up the browse tree. We assume a starting point (the project's
'   starting point), and then we load all directories from there on down.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub LoadTree(strStartPath As String)
 
    Dim oFileSystem As New FileSystemObject
    Dim oFolder As Folder
    Dim oNode As Node
    
    ' Start over.
    nTreeNodeCount = 0
    tvProjectList.Nodes.Clear
    
    ' This may not exist because it's been deleted or something. Just be sure that we
    ' have some node to process.
    If oFileSystem.FolderExists(strStartPath) Then
        ' This will be a variable in the future
        Set oFolder = oFileSystem.GetFolder(strStartPath)
        ' Top node.
        Set oNode = tvProjectList.Nodes.Add(, , oFolder.Path, oFolder.Path, "Project")
        nTreeNodeCount = nTreeNodeCount + 1
        On Error GoTo LoadFailure
        RecursiveLoad oFolder
    Else
        ' just put in a place holder.
        Set oFolder = oFileSystem.GetFolder(App.Path)
        ' Top node.
        Set oNode = tvProjectList.Nodes.Add(, , oFolder.Path, "(double click to set project path)", "Project")
        nTreeNodeCount = nTreeNodeCount + 1
    End If
    
    ' Auto-expand this
    If oNode.Children > 0 Then
        oNode.Expanded = True
    End If
    
    Exit Sub
    
LoadFailure:
    OutputLog.AddOutputLine "UNEXPECTED ERROR:"
    OutputLog.AddOutputLine Err.Description
End Sub
 
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: RecursiveLoad
' PURPOSE: Recursively loads the paths into the tree.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub RecursiveLoad(oFolder As Folder)
 
    Dim oSubFolder As Folder
    Dim oNode As Node
    
    For Each oSubFolder In oFolder.SubFolders
        Set oNode = tvProjectList.Nodes.Add(oFolder.Path, tvwChild, oSubFolder.Path, oSubFolder.Name, "FolderClosed")
        nTreeNodeCount = nTreeNodeCount + 1
        If nTreeNodeCount > 255 Then
            Err.Raise 5, "RecursiveLoad", "Too many sub-folders. FIDE Currently supports a maximum number of folders of 255. " & _
                "Set your script root project to a sub-directory and try again."
        End If
        RecursiveLoad oSubFolder
    Next
    
End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: OpenPort
' PURPOSE: Makes sure that the port is indeed open. Any port in a storm...
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub OpenPort()

    ' Attempt to open the port
    If Not SerialConnection.PortOpen Then
        ' Hard coded values for this moment. This will change later.
        SerialConnection.CommPort = CInt(frmOptions.SerialPort)
        SerialConnection.Settings = frmOptions.Baud & "," & frmOptions.Parity & "," & frmOptions.DataBits & "," & frmOptions.StopBits
        SerialConnection.InputLen = 0
        SerialConnection.PortOpen = True
        sbStatusBar.Panels(1).Text = "COM Port opened."
        
        tbToolBar.Buttons("SerialReceive").Value = tbrPressed
        tbToolBar.Buttons("SerialReceive").ToolTipText = "Serial port is open for read/write. Click to close."
    End If
    
    
End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: ClosePort
' PURPOSE: Make sure it's closed
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub ClosePort()

    ' Attempt to open the port
    If SerialConnection.PortOpen Then
        SerialConnection.PortOpen = False
        sbStatusBar.Panels(1).Text = "COM Port closed."
        
        tbToolBar.Buttons("SerialReceive").Value = tbrUnpressed
        tbToolBar.Buttons("SerialReceive").ToolTipText = "Serial port is closed. Click to close."
        
    End If
    
End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuFilePrint_Click
' PURPOSE: Print the current source file
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuFilePrint_Click()
On Error GoTo Hell

    ' The CommonDialog control is named "dlgPrint."
     
    dlgPrint.Flags = cdlPDReturnDC + cdlPDNoPageNums
    If txtActiveFile.SelLength = 0 Then
       dlgPrint.Flags = dlgPrint.Flags + cdlPDAllPages
    Else
       dlgPrint.Flags = dlgPrint.Flags + cdlPDSelection
    End If
    dlgPrint.ShowPrinter
    txtActiveFile.SelPrint dlgPrint.hDC
    Exit Sub
    
Hell:
    OutputLog.AddOutputLine "SYSTEM ERROR: "
    OutputLog.AddOutputLine Err.Number & ": " & Err.Description
    
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuHelpView9LC_Click
' PURPOSE: Go to the index help page.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuHelpView9LC_Click()
    HelpOutput.SetTopic "CompilerIndex.htm"
End Sub



' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuScriptColors_Click
' PURPOSE: Use color syntax highlighting?
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuScriptColors_Click()
    mnuScriptColors.Checked = Not mnuScriptColors.Checked
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuSearchForTopic_Click
' PURPOSE: Try to find a topic on the selected line
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuSearchForTopic_Click()
    Dim nLineCharIndex As Integer
    Dim nLineLength As Integer
    
    nLineCharIndex = SendMessage(txtActiveFile.hWnd, EM_LINEINDEX, -1&, 0&)
    nLineLength = SendMessage(txtActiveFile.hWnd, EM_LINELENGTH, -1&, 0&)
    
    HelpOutput.ScanContent Mid(txtActiveFile.Text, nLineCharIndex + 1, nLineLength)
End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuScriptPromptBefore_Click
' PURPOSE: Turn on/off prompting
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuScriptPromptBefore_Click()
    mnuScriptPromptBefore.Checked = Not mnuScriptPromptBefore.Checked
End Sub



' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuToolsROMDiff_Click
' PURPOSE: Look at two zip files. Given the two zip files (ROM sets)
'   attempt to figure out a way to know which ROM is installed.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuToolsROMDiff_Click()
On Error GoTo Hell

    ' Pop up our dialog
    Dim strCode As String
    strCode = frmROMDiff.GenerateCode(Me)
    
    ' Copy code to current window
    If Len(strCode) > 0 Then
        Dim nOriginalSel As Integer
        nOriginalSel = txtActiveFile.SelStart
        txtActiveFile.SelText = strCode
        
        If mnuScriptColors.Checked Then
            ColorSyntax.CheckRange txtActiveFile, nOriginalSel, txtActiveFile.SelStart
        End If
    
    End If
    Exit Sub
    
Hell:
    OutputLog.AddOutputLine "SYSTEM ERROR: "
    OutputLog.AddOutputLine Err.Number & ": " & Err.Description
        
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuToolsCalc_Click
' PURPOSE: Helps give you values based on bits.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuToolsCalc_Click()
     frmBinaryToHexCalc.Show , Me
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuToolsROMSignature_Click
' PURPOSE: Allows the user to add in ROM check code.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuToolsROMSignature_Click()
On Error GoTo Hell

    ' Pop up our dialog
    Dim strCode As String
    strCode = frmSignature.GenerateCode(Me)
    
    ' Copy code to current window
    If Len(strCode) > 0 Then
        Dim nOriginalSel As Integer
        nOriginalSel = txtActiveFile.SelStart
        txtActiveFile.SelText = strCode
        
        If mnuScriptColors.Checked Then
            ColorSyntax.CheckRange txtActiveFile, nOriginalSel, txtActiveFile.SelStart
        End If
    
    End If
    Exit Sub
    
Hell:
    OutputLog.AddOutputLine "SYSTEM ERROR: "
    OutputLog.AddOutputLine Err.Number & ": " & Err.Description
        
End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuToolsETERM_Click
' PURPOSE: Used to load up a compiled program, like a 6809 program.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuToolsETERM_Click()
On Error GoTo Hell

    Dim strCode As String
    strCode = frmInsertListingFile.GenerateCode(Me)
    
    ' Copy code to current window
    If Len(strCode) > 0 Then
        Dim nOriginalSel As Integer
        nOriginalSel = txtActiveFile.SelStart
        txtActiveFile.SelText = strCode
        
        If mnuScriptColors.Checked Then
            ColorSyntax.CheckRange txtActiveFile, nOriginalSel, txtActiveFile.SelStart
        End If
    
    End If
    Exit Sub
    
Hell:
    OutputLog.AddOutputLine "SYSTEM ERROR: "
    OutputLog.AddOutputLine Err.Number & ": " & Err.Description
    
End Sub



' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuViewDyanmicHelp_Click
' PURPOSE: Turn on/off the dynamic help
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuViewDyanmicHelp_Click()
    mnuViewDyanmicHelp.Checked = Not mnuViewDyanmicHelp.Checked
End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: ReceiverTimer_Timer
' PURPOSE: Check to see if we have mail...
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub ReceiverTimer_Timer()
    If SerialConnection.PortOpen Then
        Dim strBuffer As String
        strBuffer = SerialConnection.Input
        If Len(strBuffer) > 0 Then
            SerialOutputLog.AddOutput strBuffer
        End If
    End If
End Sub




' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: tsFiles_Click
' PURPOSE: The tab strip has had a selection made. We need to show that
'   file.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub tsFiles_Click()
    
    If UIFileReload Then
        UIFileReload = False
        
    Else
    
        If Not (PreviousActiveFileEditor Is Nothing) Then
            PreviousActiveFileEditor.SaveState txtActiveFile
        End If
    
        ActiveFileEditor.RestoreState txtActiveFile
        Set PreviousActiveFileEditor = ActiveFileEditor
    End If
    
    ' Change our window title
    ResetTabTitle
    
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: lvFIles_Click
' PURPOSE: A specific file has been clicked. Load it.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub lvFiles_Click()

    If (lvFiles.SelectedItem Is Nothing) Then
        Exit Sub
    End If
    
    ' Save out state
    If Not (PreviousActiveFileEditor Is Nothing) Then
        PreviousActiveFileEditor.SaveState txtActiveFile
    End If
    
    
    ' First time that we are loading something....
    If (Not tsFiles.Visible) Then
        
        tsOutput.Visible = True
        tsFiles.Visible = True
        tsFiles.Tabs.Clear
        txtActiveFile.Visible = True
        
        ' We can also enable our file operations
        ' These menu items are no longer valid
        ToggleFileMenu True

    End If
    
    ' See if this file is already loaded.
    Dim i As Integer
    Dim strFileAndDir As String
    strFileAndDir = SelectedDirectory & "\" & lvFiles.SelectedItem.Text
    For i = 1 To tsFiles.Tabs.Count
        If (tsFiles.Tabs.Item(i).Key = strFileAndDir) Then
            tsFiles.Tabs.Item(i).Selected = True
            Exit Sub
        End If
    Next
    
    Screen.MousePointer = vbHourglass
'    LockWindowUpdate txtActiveFile.hwnd

    ' If we have a previous window, save its state.
    If Not (ActiveFileEditor Is Nothing) Then
        ActiveFileEditor.SaveState txtActiveFile
    End If
    UIFileReload = True
    
    ' Create our file editor object.
    Dim obj As New FileEditor
    obj.Load txtActiveFile, lvFiles.SelectedItem.Text, SelectedDirectory
    ' So that we can save state when tabs are switched
    Set PreviousActiveFileEditor = obj

    ' Add it to the list of objects that we are tracking.
    ActiveFileCollection.Add obj, SelectedDirectory & "\" & lvFiles.SelectedItem.Text
    obj.CheckEntireDoc
    
    ' Add this new item, and select it.
    tsFiles.Tabs.Add(, SelectedDirectory & "\" & lvFiles.SelectedItem.Text, obj.WindowTitle).Selected = True
    
    ' Enable the close button
    tbTabToolBar.Buttons(1).Enabled = True
    
'    LockWindowUpdate 0
    Screen.MousePointer = vbDefault


End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: tsOutput_Click
' PURPOSE: We're changing tabs. Show the right output.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub tsOutput_Click()

    If OutputMinimized Then
        OutputLog.HideOutput
        SerialOutputLog.HideOutput
        HelpOutput.HideOutput
    Else
    
        If tsOutput.SelectedItem = OutputLog.Caption Then
            OutputLog.ShowOutput
        ElseIf tsOutput.SelectedItem = SerialOutputLog.Caption Then
            SerialOutputLog.ShowOutput
        Else
            HelpOutput.ShowOutput
        End If
    
        EnableTabButtons
        
    End If
        
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: EnableTabButtons
' PURPOSE: Turns on/off the smaller tab buttons
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub EnableTabButtons()
    tbOutputTabs.Buttons("SaveOutput").Enabled = txtSerialOutput.Visible Or txtOutput.Visible
    tbOutputTabs.Buttons("ClearOuput").Enabled = txtSerialOutput.Visible Or txtOutput.Visible
End Sub



' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: txtActiveFile_Change
' PURPOSE: The contents have changed. Make sure that the FileEditor class
'   knows about this.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub txtActiveFile_Change()

    If UIFileReload Or LockEditorWindow Then
        ' we're in the middle of a file reload process. skip processing
        Exit Sub
    End If

    UIFileDirty = True
        
    Dim bOriginalState As Boolean
    
    bOriginalState = ActiveFileEditor.FileHasChanged
    ActiveFileEditor.UpdateTextContents txtActiveFile
    
    ' If we go from unchanged to changed, make text updates
    If bOriginalState <> ActiveFileEditor.FileHasChanged Then
        ResetTabTitle
    End If
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: txtActiveFile_KeyDown
' PURPOSE: Used to prevent double-pasting.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub txtActiveFile_KeyDown(KeyCode As Integer, Shift As Integer)
   
    ' A control key
    If Shift = vbCtrlMask Then
    
        '67 = ^C
        '86 = ^V
        '88 = ^X
        If KeyCode = 67 Then
            mnuEditCopy_Click
            KeyCode = 0
            Exit Sub
        ElseIf KeyCode = 86 Then
            mnuEditPaste_Click
            KeyCode = 0
            Exit Sub
        ElseIf KeyCode = 88 Then
            mnuEditCut_Click
            KeyCode = 0
            Exit Sub
        End If
    End If
    
    ' An ALT key
    If Shift = vbAltMask Then
        If KeyCode = 8 Then
            ' We have the same as UNDO
            mnuEditUndo_Click
            KeyCode = 0
            Exit Sub
        End If
    End If
   
End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: ResetTabTitle
' PURPOSE: Sets the title & status bar
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub ResetTabTitle()

    ActiveTab.Caption = ActiveFileEditor.WindowTitle
    If (ActiveFileEditor.FileHasChanged) Then
        Me.Caption = "FIDE: '" & ActiveFileEditor.FileNameAndDirectory & "' *"
        sbStatusBar.Panels(4).Text = "MODIFIED"
    Else
        Me.Caption = "FIDE: '" & ActiveFileEditor.FileNameAndDirectory & "'"
        sbStatusBar.Panels(4).Text = ""
    End If

End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: UpdateCursorStats
' PURPOSE: Update cursor position
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub txtActiveFile_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UpdateCursorStats
End Sub
Private Sub txtActiveFile_GotFocus()
    UpdateCursorStats
End Sub
Private Sub txtActiveFile_KeyUp(KeyCode As Integer, Shift As Integer)
    UpdateCursorStats
    If UIFileDirty Then
        ActiveFileEditor.CheckEditedText
        UIFileDirty = False
    End If
End Sub



' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: UpdateCursorStats
' PURPOSE: Tell the user where they left their cursor
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub UpdateCursorStats()
    Dim nLineNumber As Integer
    Dim nLineCharIndex As Integer
    Dim nCurrCol As Integer
    Dim nLineLength As Integer
    
    nLineNumber = SendMessage(txtActiveFile.hWnd, EM_LINEFROMCHAR, -1&, 0&)
    nLineCharIndex = SendMessage(txtActiveFile.hWnd, EM_LINEINDEX, -1&, 0&)
    nLineLength = SendMessage(txtActiveFile.hWnd, EM_LINELENGTH, -1&, 0&)
    
    nCurrCol = txtActiveFile.SelStart - nLineCharIndex + 1
    sbStatusBar.Panels(2).Text = "Line " & (nLineNumber + 1)
    sbStatusBar.Panels(3).Text = "Col " & nCurrCol
    
    ' Don't send to the HelpOutput unless we're being asked to.
    If mnuViewDyanmicHelp.Checked Then
        HelpOutput.ScanContent Mid(txtActiveFile.Text, nLineCharIndex + 1, nLineLength)
    End If
End Sub



' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: tvProjectList_Expand
' PURPOSE: Set the appropriate icon for the node
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub tvProjectList_Expand(ByVal Node As MSComctlLib.Node)
    If Node.Index <> 1 Then
        Node.Image = "FolderOpen"
    End If
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: tvProjectList_Collapse
' PURPOSE: Set the appropriate icon for the node
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub tvProjectList_Collapse(ByVal Node As MSComctlLib.Node)
    If Node.Index <> 1 Then
        Node.Image = "FolderClosed"
    Else
        Node.Expanded = True
    End If
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: tvProjectList_Click
' PURPOSE: Has our selection changed?
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub tvProjectList_Click()
On Error GoTo Hell

    If tvProjectList.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    If tvProjectList.SelectedItem.Index = 1 Then
        tbToolBar.Buttons("DeleteFolder").Enabled = False
        tbToolBar.Buttons("EditFolder").Enabled = False
        mnuProjectDeleteFolder.Enabled = False
        mnuProjectEditName.Enabled = False
    Else
        tbToolBar.Buttons("DeleteFolder").Enabled = True
        tbToolBar.Buttons("EditFolder").Enabled = True
        mnuProjectDeleteFolder.Enabled = True
        mnuProjectEditName.Enabled = True
    End If
    
    ' Let the user know what our directory is
    sbStatusBar.Panels(1).Text = "Reading " & SelectedDirectory & " directory..."
    

    ' Open up this directory so that we can show the files
    Dim oFileSystem As New FileSystemObject
    Dim oFolder As Folder
    Dim oFile As file
    Dim oListItem As ListItem
    
    Set oFolder = oFileSystem.GetFolder(SelectedDirectory)
    lvFiles.ListItems.Clear
    Dim strArray
    Dim strExtension As String

    ' Only a few files we will actually give access to:
    For Each oFile In oFolder.Files
        strArray = Split(oFile.Name, ".")
        strExtension = LCase(strArray(UBound(strArray)))
        Select Case (strExtension)
            Case "s", "9lc"
                Set oListItem = lvFiles.ListItems.Add(, oFile.Name, oFile.Name, , "Script")
                oListItem.ToolTipText = oFile.Name & " (Script file)"
                oListItem.SubItems(1) = "B" & oFile.Name
            Case "txt"
                Set oListItem = lvFiles.ListItems.Add(, oFile.Name, oFile.Name, , "Notes")
                oListItem.ToolTipText = oFile.Name & " (Text/Notes file)"
                oListItem.SubItems(1) = "A" & oFile.Name
                
            ' Show them, but grey them out
            Case Else
                Set oListItem = lvFiles.ListItems.Add(, oFile.Name, oFile.Name)
                oListItem.ForeColor = &HC0C0C0
                oListItem.ToolTipText = "Cannot be viewed in FIDE"
                oListItem.SubItems(1) = "Z" & oFile.Name
        End Select
        
        
    Next
    sbStatusBar.Panels(1).Text = sbStatusBar.Panels(1).Text & "done."
    Exit Sub
    
Hell:
    OutputLog.AddOutputLine "UNEXPECTED ERROR:"
    OutputLog.AddOutputLine Err.Description
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: tvProjectList_DblClick
' PURPOSE: Will lauch the project settings.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub tvProjectList_DblClick()
    If tvProjectList.SelectedItem.Index = 1 Then
        mnuProjectSetRoot_Click
    End If
End Sub




' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: imgHSplitter_MouseDown
' PURPOSE: The horizontal splitter is about to be moved.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub imgHSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With imgHSplitter
        picHSplitter.Move .Left \ 2, .Top - 20, .Width, picHSplitter.Height
    End With
    picHSplitter.Visible = True
    mbHMoving = True
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: imgHSplitter_MouseMove
' PURPOSE: The mouse is being moved with the splitter selected
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub imgHSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglPos As Single

    If mbHMoving Then
        sglPos = Y + imgHSplitter.Top
        If sglPos < sglHSplitLimit Then
            picHSplitter.Top = sglHSplitLimit
        ElseIf sglPos > Me.Height - sglHSplitLimit Then
            picHSplitter.Top = Me.Height - sglHSplitLimit
        Else
            picHSplitter.Top = sglPos
        End If
    End If
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: imgHSplitter_MouseDown
' PURPOSE: We're done moving the horizontal splitter
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub imgHSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SizeControls
    picHSplitter.Visible = False
    mbHMoving = False
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: imgH2Splitter_MouseDown
' PURPOSE: The horizontal splitter is about to be moved.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub imgH2Splitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With imgH2Splitter
        picH2Splitter.Move picSplitter.Left + 40, .Top - 20, .Width, picH2Splitter.Height
    End With
    picH2Splitter.Visible = True
    mbH2Moving = True
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: imgH2Splitter_MouseMove
' PURPOSE: The mouse is being moved with the splitter selected
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub imgH2Splitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglPos As Single

    If mbH2Moving Then
        sglPos = Y + imgH2Splitter.Top
        If sglPos < sglH2SplitLimit Then
            picH2Splitter.Top = sglH2SplitLimit
        ElseIf sglPos > Me.Height - sglH2SplitLimit Then
            picH2Splitter.Top = Me.Height - sglH2SplitLimit
        Else
            picH2Splitter.Top = sglPos
        End If
    End If
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: imgH2Splitter_MouseDown
' PURPOSE: We're done moving the horizontal splitter
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub imgH2Splitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SizeControls
    picH2Splitter.Visible = False
    mbH2Moving = False
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: imgSplitter_MouseDown
' PURPOSE: The vertical splitter is about to be moved.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With imgSplitter
        picSplitter.Move .Left \ 2, .Top - 20, picSplitter.Width, .Height
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: imgSplitter_MouseMove
' PURPOSE: The mouse is being moved with the splitter selected
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglPos As Single

    If mbMoving Then
        sglPos = X + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: imgSplitter_MouseDown
' PURPOSE: We're done moving the vertical splitter
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SizeControls
    picSplitter.Visible = False
    mbMoving = False
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: SizeControls
' PURPOSE: Set up parameters to pass into ReSizeControls
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub SizeControls()
    ReSizeControls picSplitter.Left, picHSplitter.Top, picH2Splitter.Top
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: SizeControls
' PURPOSE: Invoked from Form_Resize and used to resize our controls
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub ReSizeControls(X As Single, Y As Single, Y2 As Single)
    On Error Resume Next

    ' Limits on dimensions
    If X < 1500 Then X = 1500
    If X > (Me.Width - 1500) Then X = Me.Width - 1500
    
    If Y < 800 Then Y = 800
    If Y > (Me.Height - 800) Then Y = Me.Height - 800
    
    If Y2 < 1600 Then Y2 = 1600
    If Y2 > (Me.Height - 2100) Then Y2 = Me.Height - 2100
    If OutputMinimized Then Y2 = Me.Height - 1700
    
    ' Right aligned part of our tool bar
    tbTabToolBar.Left = Me.Width - tbTabToolBar.Width - 100
    
    ' Get the height of our workspace.
    Dim nHeight
    If sbStatusBar.Visible Then
        ' The other control that we need to add will use this
        nHeight = Me.ScaleHeight - (sbStatusBar.Height)
    Else
        nHeight = Me.ScaleHeight
    End If
    
    ' Progress bar
    pbSend.Left = sbStatusBar.Left
    pbSend.Top = sbStatusBar.Top
    pbSend.Width = sbStatusBar.Width
    pbSend.Height = sbStatusBar.Height
    
    imgSplitter.Left = X
    imgSplitter.Top = 0
    If tbToolBar.Visible Then imgSplitter.Top = tbToolBar.Height
    imgSplitter.Height = nHeight
    
    imgHSplitter.Top = Y
    imgHSplitter.Width = X
    
    imgH2Splitter.Top = Y2
    imgH2Splitter.Left = X + 40
    imgH2Splitter.Width = Me.Width - X
        
    ' Set the project tree's dimensions
    tvProjectList.Width = imgSplitter.Left
    tvProjectList.Top = 0
    If tbToolBar.Visible Then tvProjectList.Top = tbToolBar.Height
    tvProjectList.Height = imgHSplitter.Top - tvProjectList.Top
    
    
    ' Set the contents of that node (the files)
    lvFiles.Width = imgSplitter.Left
    lvFiles.Top = imgHSplitter.Top + imgHSplitter.Height
    lvFiles.Height = nHeight - (imgHSplitter.Top + imgHSplitter.Height)
    
    ' if we have one or more open files.
    'If (tsFiles.Visible) Then  ' jsb: just always resize the controls here
        tsFiles.Left = imgSplitter.Left + imgSplitter.Width
        tsFiles.Width = Me.Width - (tsFiles.Left + 150)
        tsFiles.Top = 450
        tsFiles.Height = Y2 - tsFiles.Top
        
        ' Our text box
        txtActiveFile.Left = tsFiles.Left + 80
        txtActiveFile.Width = tsFiles.Width - 180
        txtActiveFile.Top = tsFiles.Top + 400
        txtActiveFile.Height = tsFiles.Height - 460
        
    'End If
    
    ' Second set of output tabs
    'If (tsOutput.Visible) Then
    
        ' Our title area
        tsOutput.Left = imgSplitter.Left + imgSplitter.Width
        tsOutput.Width = Me.Width - (tsOutput.Left + 150)
        tsOutput.Top = imgH2Splitter.Top + imgH2Splitter.Height
        tsOutput.Height = (Me.ScaleHeight) - (tsOutput.Top + 300)
        
        
        tbOutputTabs.Top = imgH2Splitter.Top
        tbOutputTabs.Left = (tsOutput.Left + tsOutput.Width) - tbOutputTabs.Width
        
        ' If we have our output visible
        txtOutput.Left = tsOutput.Left + 60
        txtOutput.Width = tsOutput.Width - 180
        txtOutput.Top = tsOutput.Top + 60
        txtOutput.Height = tsOutput.Height - 400
        
        ' This is the serial output section.
        txtSerialOutput.Left = tsOutput.Left + 60
        txtSerialOutput.Width = tsOutput.Width - 180
        txtSerialOutput.Top = tsOutput.Top + 60
        txtSerialOutput.Height = tsOutput.Height - 400
        ' Browser controls
        wbHelpContext.Left = tsOutput.Left + 60
        wbHelpContext.Width = tsOutput.Width - 180
        wbHelpContext.Top = tsOutput.Top + 60
        wbHelpContext.Height = tsOutput.Height - 400
        
        EnableTabButtons
    'End If
    
    txtActiveFile.RightMargin = 2 * txtActiveFile.Width

 
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: tbOutputTabs_ButtonClick
' PURPOSE: Special small button commands.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub tbOutputTabs_ButtonClick(ByVal Button As MSComctlLib.Button)
    OutputTabCommand Button.Key
End Sub
Private Sub OutputTabCommand(strCommand As String)
    On Error Resume Next
    Select Case strCommand
            
        ' First part of the button bar
        Case "MinimizeOutput"
            OutputMinimized = True
            nLastSplitterPos = picH2Splitter.Top
            ReSizeControls picSplitter.Left, picHSplitter.Top, Me.Height
            tsOutput_Click
        Case "MaximizeOutput"
            If OutputMinimized Then
                OutputMinimized = False
                ReSizeControls picSplitter.Left, picHSplitter.Top, nLastSplitterPos
                tsOutput_Click
            Else
                ReSizeControls picSplitter.Left, picHSplitter.Top, 0
            End If
        Case "MaximizeForCompile"
            If OutputMinimized Then
                OutputMinimized = False
                ReSizeControls picSplitter.Left, picHSplitter.Top, nLastSplitterPos
            End If
        Case "ClearOuput"
            If txtSerialOutput.Visible Then
                SerialOutputLog.ClearOutput
            ElseIf txtOutput.Visible Then
                OutputLog.ClearOutput
            End If
        Case "SaveOutput"
            SaveOutputFile
        
    End Select

End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: SaveOutputFile
' PURPOSE: Saves whatever we have in the log file.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub SaveOutputFile()
    Dim sFile As String

    With dlgCommonDialog
        .InitDir = SelectedDirectory
        .FileName = ""
        .DialogTitle = "Save Output File"
        .CancelError = False
        .Filter = "Text File (*.txt)|*.txt|All Files (*.*)|*.*"
        .ShowSave
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    
    ' Get the content we want to save.
    Dim strContents As String
    If txtSerialOutput.Visible Then
        strContents = txtSerialOutput.Text
    ElseIf txtOutput.Visible Then
        strContents = txtOutput.Text
    End If
    
    ' Give it at least one character so that
    ' we don't have an error
    If Len(strContents) = 0 Then
        strContents = strContents & " "
    End If
    
    On Error GoTo CantSaveLog
    ' Save it.
    Dim oFileSystem As New FileSystemObject
    Dim oTextStream As TextStream
    Set oTextStream = oFileSystem.CreateTextFile(sFile)
    oTextStream.Write strContents
    oTextStream.Close
    Exit Sub
    
CantSaveLog:
    MsgBox "ERROR: Could not save the output file. Make sure you have " & vbCrLf & _
        "local file write access.", vbExclamation, "ERROR: could not save log"
        
        
    
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: tbToolBar_ButtonClick
' PURPOSE: Process commands from the tool bar
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    On Error Resume Next
    Select Case Button.Key
            
        ' First part of the button bar
        Case "EditFolder"
            mnuProjectEditName_Click
        Case "NewFolder"
            mnuProjectAddFolder_Click
        Case "DeleteFolder"
            mnuProjectDeleteFolder_Click
            
        ' Second part -- new and save
        Case "FileNew"
            mnuFileNew_Click
        Case "FileSave"
            mnuFileSave_Click
        
        ' Third part -- cut, copy, and past
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
            
        ' Fourth, compile
        Case "Compile"
            mnuScriptCompile_Click
        
        ' Fifth, serial options
        Case "Properties"
            mnuFileProperties_Click
        Case "SerialReceive"
            mnuScriptOpen_Click
        Case "SerialSend"
            mnuScriptCompileSend_Click
        Case "BinarySend"
            mnuScriptBinarySend_Click
        
            
    End Select
End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuProjectSetRoot_Click
' PURPOSE: We may be trying to reset out root project.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuProjectSetRoot_Click()
On Error GoTo Hell

    If frmRootLocation.ShowModal(Me) Then
        ' We'll need to re-load the tree to reflect our changes.
        LoadTree frmRootLocation.ProjectPath
    End If
    Exit Sub
    
Hell:
    OutputLog.AddOutputLine "SYSTEM ERROR: "
    OutputLog.AddOutputLine Err.Number & ": " & Err.Description
        
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuProjectEditName_Click
' PURPOSE: rename this folder.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuProjectEditName_Click()
On Error GoTo Hell

    Dim strNewFolderName As String
    Dim strCurrentFolder As String
    Dim strNewPath As String
    Dim arrPathBreakDown
    
    If Len(SelectedDirectory) = 0 Then
        Exit Sub
    End If
    
    arrPathBreakDown = Split(SelectedDirectory, "\")
    strCurrentFolder = arrPathBreakDown(UBound(arrPathBreakDown))
    
    dlgRenameFolder.NewFolderName = strCurrentFolder
    strNewFolderName = dlgRenameFolder.GetNewFolder(Me)
    
    ' We have a real folder to try and add.
    If Len(strNewFolderName) > 0 Then
        Dim oFileSystem As New FileSystemObject
        On Error GoTo FolderCouldNotBeCreated
        strNewPath = Replace(SelectedDirectory, "\" & strCurrentFolder, "\" & strNewFolderName)
        oFileSystem.MoveFolder SelectedDirectory, strNewPath
        
        On Error GoTo RefreshError
        LoadTree frmRootLocation.ProjectPath
    End If
    Exit Sub

FolderCouldNotBeCreated:
    MsgBox "Folder could not be renamed. Check file name", vbOKOnly, "ERROR"
    Exit Sub
    
RefreshError:
    OutputLog.AddOutputLine "ERROR: Could not refresh project view: " & frmRootLocation.ProjectPath & "."
    OutputLog.AddOutputLine "Make sure that directory still exists."
    Exit Sub
    
Hell:
    OutputLog.AddOutputLine "SYSTEM ERROR: "
    OutputLog.AddOutputLine Err.Number & ": " & Err.Description
        
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuProjectAddFolder_Click
' PURPOSE: We're adding a new folder
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuProjectAddFolder_Click()
On Error GoTo Hell

    Dim strNewFolderName As String
    strNewFolderName = dlgAddNewFolder.GetNewFolder(Me)
    
    ' We have a real folder to try and add.
    If Len(strNewFolderName) > 0 Then
        Dim oFileSystem As New FileSystemObject
        On Error GoTo FolderCouldNotBeCreated
        
        oFileSystem.CreateFolder SelectedDirectory & "\" & strNewFolderName
        
        On Error GoTo RefreshError
        LoadTree frmRootLocation.ProjectPath
                
    End If
    Exit Sub

FolderCouldNotBeCreated:
    MsgBox "Folder could not be created. Check file name", vbOKOnly, "ERROR"
    Exit Sub
    
RefreshError:
    OutputLog.AddOutputLine "ERROR! Could not refresh project view: " & frmRootLocation.ProjectPath & "."
    OutputLog.AddOutputLine "Make sure that directory still exists."
    Exit Sub
    
Hell:
    OutputLog.AddOutputLine "SYSTEM ERROR: "
    OutputLog.AddOutputLine Err.Number & ": " & Err.Description
        
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuProjectDeleteFolder_Click
' PURPOSE: Dangerous! Do you really want to delete this folder!?!
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuProjectDeleteFolder_Click()
On Error GoTo Hell

    Dim strFolderToDelete As String
    Dim strCurrentFolder As String
    Dim arrPathBreakDown
    
    ' Check for error
    If Len(SelectedDirectory) = 0 Then
        Exit Sub
    End If
    
    arrPathBreakDown = Split(SelectedDirectory, "\")
    strCurrentFolder = arrPathBreakDown(UBound(arrPathBreakDown))
        
    Dim answer As VbMsgBoxResult
    answer = MsgBox("This will permanently delete " & vbCrLf & vbCrLf & vbTab & SelectedDirectory & "." & vbCrLf & vbCrLf & "Are you sure?", vbYesNo, "Are you sure?")
    
    
    ' Can means go away
    If answer = vbNo Then
        Exit Sub
    End If
    
    ' Okay, let's rock
    On Error GoTo FolderCouldNotBeDeleted
    Dim strNewPath As String
    Dim strNewLocation As String
    
    strNewPath = Replace(SelectedDirectory, "\" & strCurrentFolder, "")
    strNewLocation = GetSetting(App.Title, "Settings", "Trash", App.Path & "\Trash")
    
    strFolderToDelete = SelectedDirectory
    Dim oFileSystem As New FileSystemObject
    
    ' I think we're just going to move this folder to the trash location... for now
    ' oFileSystem.DeleteFolder strFolderToDelete, True
    If Not oFileSystem.FolderExists(strNewLocation) Then
        oFileSystem.CreateFolder strNewLocation
    End If
    strNewLocation = strNewLocation & "\" & strCurrentFolder
    
    ' Move it to the trash.
    oFileSystem.MoveFolder strFolderToDelete, strNewLocation
    
    On Error GoTo RefreshError
    LoadTree frmRootLocation.ProjectPath
    
    Exit Sub

FolderCouldNotBeDeleted:
    MsgBox "Folder could not be deleted. Check folder and make sure it is empty.", vbOKOnly, "ERROR"
    Exit Sub
    
RefreshError:
    OutputLog.AddOutputLine "ERROR! Could not refresh project view: " & frmRootLocation.ProjectPath & "."
    OutputLog.AddOutputLine "Make sure that directory still exists."
    Exit Sub
    
Hell:
    OutputLog.AddOutputLine "SYSTEM ERROR: "
    OutputLog.AddOutputLine Err.Number & ": " & Err.Description

End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: tbTabToolBar_ButtonClick
' PURPOSE: Sub tool bar that handles a few other function
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub tbTabToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "CloseCurrentFile"
            mnuFileClose_Click
    End Select
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuFileClose_Click
' PURPOSE: Close the current file that we have selected.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuFileClose_Click()
On Error GoTo Hell

    ' check for saving....
    If Not (ActiveFileEditor.CanClose) Then
        Exit Sub
    End If
    
    Dim nNextActiveIndex As Integer
    ActiveFileCollection.Remove (ActiveTab.Key)
    nNextActiveIndex = tsFiles.SelectedItem.Index
    tsFiles.Tabs.Remove (tsFiles.SelectedItem.Index)
    
    
    ' If we don't have any tabs left, then don't show junk...
    If tsFiles.Tabs.Count < 1 Then
        tsFiles.Visible = False
        txtActiveFile.Visible = False
        ' Disable the close button
        tbTabToolBar.Buttons(1).Enabled = False
        
        ' These menu items are no longer valid
        ToggleFileMenu False
        
    Else
        nNextActiveIndex = nNextActiveIndex - 1
        If nNextActiveIndex > 0 Then
            tsFiles.Tabs(nNextActiveIndex).Selected = True
        Else
            tsFiles_Click
        End If
    End If
    Exit Sub
    
Hell:
    OutputLog.AddOutputLine "SYSTEM ERROR: "
    OutputLog.AddOutputLine Err.Number & ": " & Err.Description
    
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuScriptOpen_Click
' PURPOSE: Make sure the serial port is open and ready to receive commands
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuScriptOpen_Click()
On Error GoTo Hell


    If SerialConnection.PortOpen Then
        ClosePort
    Else
        OpenPort    ' That's it...
    End If
    
    Exit Sub
    
Hell:
    OutputLog.AddOutputLine "SYSTEM ERROR: "
    OutputLog.AddOutputLine Err.Number & ": " & Err.Description
        
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuScriptBinarySend_Click
' PURPOSE: Send a file, without compiling it
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuScriptBinarySend_Click()
On Error GoTo Hell

    ' Prompt user to prep the unit
    If Not (frmPrepare9010.ShowPrepare(Me) = vbOK) Then
        Exit Sub
    End If
    
    ' Compile, and then send.
    If Not (ActiveFileEditor.OnSave) Then
        Exit Sub
    End If
    
    ' Finally, send it away
    TransmitFileto9010A ActiveFileEditor.FileNameAndDirectory
    Exit Sub
    
Hell:
    OutputLog.AddOutputLine "SYSTEM ERROR: "
    OutputLog.AddOutputLine Err.Number & ": " & Err.Description
        
    

End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuScriptCompileSend_Click
' PURPOSE: Compile a file, then send it.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuScriptCompileSend_Click()
On Error GoTo Hell

    ' Prompt user to prep the unit
    If Not (frmPrepare9010.ShowPrepare(Me) = vbOK) Then
        Exit Sub
    End If
    
    ' Compile, and then send.
    If Not (ActiveFileEditor.OnSave) Then
        Exit Sub
    End If
    
    OutputTabCommand "MaximizeForCompile"
    If Not (Compiler.Compile) Then
        Exit Sub
    End If
    
    ' Finally, send it away
    TransmitFileto9010A Compiler.CompiledDirectoryAndFileName
    Exit Sub
    
Hell:
    OutputLog.AddOutputLine "SYSTEM ERROR: "
    OutputLog.AddOutputLine Err.Number & ": " & Err.Description
        
    
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: TransmitFileto9010A
' PURPOSE: Sends this named file to the 9010A. Should have already
'   been compiled
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub TransmitFileto9010A(strFileName As String)

    Dim oFileSystem As New FileSystemObject
    Dim oTextStream As TextStream
    Set oTextStream = oFileSystem.OpenTextFile(strFileName, ForReading)
    Dim strLine As String
    
    Dim nLineCount As Integer
    nLineCount = 0
    While Not oTextStream.AtEndOfStream
        oTextStream.ReadLine
        nLineCount = nLineCount + 1
    Wend
    oTextStream.Close
    
   ' This is the moment we've been waiting for...
    OpenPort
 
    ' Feed it line by line and wait for each line to be consumed
    pbSend.Visible = True
    pbSend.Value = 0
    pbSend.Max = nLineCount
    
    ' Open it again, this time to send.
    Set oTextStream = oFileSystem.OpenTextFile(strFileName, ForReading)
    While Not oTextStream.AtEndOfStream
        strLine = oTextStream.ReadLine
        
        SerialConnection.Output = strLine & vbCrLf

        pbSend.Value = pbSend.Value + 1
        
        While SerialConnection.OutBufferCount > 0
            DoEvents
        Wend
    Wend
    
    oTextStream.Close
    pbSend.Visible = False
    
    sbStatusBar.Panels(1).Text = "File sent to 9010A."
End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuScriptCompile_Click
' PURPOSE: Compile a file
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuScriptCompile_Click()
On Error GoTo Hell
    ' Let the compiler do the work
    If ActiveFileEditor.OnSave Then
        txtActiveFile_Change
        OutputTabCommand "MaximizeForCompile"
        Compiler.Compile
    End If
    Exit Sub
    
Hell:
    OutputLog.AddOutputLine "SYSTEM ERROR: "
    OutputLog.AddOutputLine Err.Number & ": " & Err.Description
        
End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuFileSave_Click
' PURPOSE: Save the current file
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuFileSave_Click()
On Error GoTo Hell

    ActiveFileEditor.OnSave
    ResetTabTitle
    Exit Sub
    
Hell:
    OutputLog.AddOutputLine "SYSTEM ERROR: "
    OutputLog.AddOutputLine Err.Number & ": " & Err.Description
    
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuFileSaveAll_Click
' PURPOSE: Saves all files
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuFileSaveAll_Click()
On Error GoTo Hell

    Dim i As Integer
    Dim objFileEditor As FileEditor
    
    For i = 1 To tsFiles.Tabs.Count
        
        Set objFileEditor = ActiveFileCollection(tsFiles.Tabs(i).Key)
        ' Break out if we're told not to save (it's like a cancel)
        If Not (objFileEditor.OnSave) Then
            Exit Sub
        End If
    
        ' We've saved it, we may have to update the title
        tsFiles.Tabs(i).Caption = objFileEditor.WindowTitle
    
    Next
    
    ' Change our window title
    ResetTabTitle
    Exit Sub
    
Hell:
    OutputLog.AddOutputLine "SYSTEM ERROR: "
    OutputLog.AddOutputLine Err.Number & ": " & Err.Description
    
End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuFileProperties_Click
' PURPOSE: Application properties
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuFileProperties_Click()
    frmOptions.Show vbModal, Me
    
    ' Don't ever disabled the properties for serial port
    ' tbToolBar.Buttons("Properties").Enabled = frmOptions.SerialSettingsSet
    ToggleFileMenu (tsFiles.Tabs.Count > 0)
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuFileExit_Click
' PURPOSE: Terminate
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuFileExit_Click()
On Error GoTo Hell

    Unload Me
    Exit Sub
    
Hell:
    OutputLog.AddOutputLine "SYSTEM ERROR: "
    OutputLog.AddOutputLine Err.Number & ": " & Err.Description
    
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuEditPaste_Click
' PURPOSE: paste selection
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuEditPaste_Click()
On Error GoTo Hell
    
    Dim nOriginalStart As Integer
    nOriginalStart = Screen.ActiveControl.SelStart
    Screen.ActiveControl.SelText = Clipboard.GetText()
     
    ' Might have to reformat the data...
    If Screen.ActiveControl = txtActiveFile Then
        ActiveFileEditor.CheckRange nOriginalStart, Len(Clipboard.GetText)
    End If
    
    Exit Sub
     
Hell:
    OutputLog.AddOutputLine "SYSTEM ERROR: "
    OutputLog.AddOutputLine Err.Number & ": " & Err.Description
    
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuEditCopy_Click
' PURPOSE: copy selection
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuEditCopy_Click()
On Error GoTo Hell
    Clipboard.SetText (Screen.ActiveControl.SelText)
    Exit Sub
    
Hell:
    OutputLog.AddOutputLine "SYSTEM ERROR: "
    OutputLog.AddOutputLine Err.Number & ": " & Err.Description
    
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuEditCut_Click
' PURPOSE: cut selection
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuEditCut_Click()
On Error GoTo Hell

    ' Clear the contents of the Clipboard.
    Clipboard.Clear
    ' Copy selected text to Clipboard.
    Clipboard.SetText Screen.ActiveControl.SelText
    ' Delete selected text.
    Screen.ActiveControl.SelText = ""
    Exit Sub

Hell:
    OutputLog.AddOutputLine "SYSTEM ERROR: "
    OutputLog.AddOutputLine Err.Number & ": " & Err.Description
    
End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuFileDelete_Click
' PURPOSE: Delete the current file.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuFileDelete_Click()
On Error GoTo Hell

    Dim strFileToKill As String
    strFileToKill = ActiveFileEditor.FileNameAndDirectory
    
    ' Confirm you REALLY want to do this.
    Dim answer As VbMsgBoxResult
    answer = MsgBox("Are you really sure you want to PERMANENTLY" & vbCrLf & _
            "delete this file? This CANNOT be undone." & vbCrLf & _
            vbCrLf & _
            vbTab & strFileToKill & vbCrLf, vbYesNo, "PERMANENTLY Delete " & strFileToKill)
            
    If Not (answer = vbYes) Then
        Exit Sub
    End If
    
    ActiveFileCollection.Remove (ActiveTab.Key)
    tsFiles.Tabs.Remove (tsFiles.SelectedItem.Index)
    
    ' Actually (really) delete the file now
    Dim oFileSystem As New FileSystemObject
    oFileSystem.DeleteFile strFileToKill
    
    ' If we don't have any tabs left, then don't show junk...
    If tsFiles.Tabs.Count < 1 Then
        tsFiles.Visible = False
        txtActiveFile.Visible = False
        ' Disable the close button
        tbTabToolBar.Buttons(1).Enabled = False
        ' These menu items are no longer valid
        ToggleFileMenu False
    End If

    ' Refresh out lists
    tvProjectList_Click
    If tsFiles.Tabs.Count >= 1 Then
        tsFiles_Click
    End If
    
    Exit Sub
    
Hell:
    OutputLog.AddOutputLine "SYSTEM ERROR: "
    OutputLog.AddOutputLine Err.Number & ": " & Err.Description
    
End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuFileNew_Click
' PURPOSE: We're going to make them specify the file location first.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuFileNew_Click()
On Error GoTo Hell

    Dim sFile As String

    With dlgCommonDialog
        .InitDir = SelectedDirectory
        .FileName = ""
        .DialogTitle = "New File"
        .CancelError = False
        .Filter = "9010A Script (*.9lc)|*.9lc|Script file (*.s)|*.s|Work Notes Text File (*.txt)|*.txt|All Files (*.*)|*.*"
        .ShowSave
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    
    Dim sName As String
    Dim sDirectory As String
    
    sName = Right(sFile, Len(sFile) - InStrRev(sFile, "\"))
    sDirectory = Left(sFile, InStrRev(sFile, "\"))
    
    ' See if this file is already loaded.
    Dim i As Integer
    For i = 1 To tsFiles.Tabs.Count
        If (tsFiles.Tabs.Item(i).Key = sFile) Then
            ' It is. Easiest thing is to close it first.
            tsFiles.Tabs.Item(i).Selected = True
            mnuFileClose_Click
        End If
    Next
    
    ' Now... we need to open up our file, add a tab, and so on.
    ' First time that we are loading something....
    If (Not tsFiles.Visible) Then
        tsOutput.Visible = True
        tsFiles.Visible = True
        tsFiles.Tabs.Clear
        txtActiveFile.Visible = True
        SizeControls
        
        ' These menu items are no longer valid
        ToggleFileMenu True
        
    End If
    
    ' Create our file editor object.
    Dim obj As New FileEditor
    obj.CreateNew sName, sDirectory
    
    ' Add it to the list of objects that we are tracking.
    ActiveFileCollection.Add obj, sFile
    
    tsFiles.Tabs.Add(, sFile, sName).Selected = True
    
    ' Enable the close button
    tbTabToolBar.Buttons(1).Enabled = True
    
    ' Causes a refresh to occur, in case this was added to the selected directory
    tvProjectList_Click
    Exit Sub
    
Hell:
    OutputLog.AddOutputLine "SYSTEM ERROR: "
    OutputLog.AddOutputLine Err.Number & ": " & Err.Description
    
End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuHelpAbout_Click
' PURPOSE: Show the help about dialog
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuEditSelectAll_Click
' PURPOSE: Select all available text
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuEditSelectAll_Click()
On Error GoTo Hell

    If Not (ActiveFileEditor Is Nothing) Then
        ActiveFileEditor.SelectAll txtActiveFile
    End If
    Exit Sub
    
Hell:
    OutputLog.AddOutputLine "SYSTEM ERROR: "
    OutputLog.AddOutputLine Err.Number & ": " & Err.Description
        
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: mnuEditUndo_Click
' PURPOSE: Allows the user one level of UNDO
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuEditUndo_Click()
On Error GoTo Hell

    If Not (ActiveFileEditor Is Nothing) Then
        ActiveFileEditor.DoUndo txtActiveFile
    End If
    Exit Sub
    
Hell:
    OutputLog.AddOutputLine "SYSTEM ERROR: "
    OutputLog.AddOutputLine Err.Number & ": " & Err.Description
    
End Sub

Private Sub mnuFileRename_Click()
On Error GoTo Hell

    'ToDo: Add 'mnuFileRename_Click' code.
    MsgBox "Add 'mnuFileRename_Click' code."
    Exit Sub
    
Hell:
    OutputLog.AddOutputLine "SYSTEM ERROR: "
    OutputLog.AddOutputLine Err.Number & ": " & Err.Description
    
End Sub
