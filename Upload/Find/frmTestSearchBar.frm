VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{AC551970-584E-49C8-8BE1-4EF304A6250C}#1.1#0"; "vbalExpBar6.ocx"
Begin VB.Form frmFind 
   Caption         =   "Find Files"
   ClientHeight    =   6990
   ClientLeft      =   2355
   ClientTop       =   2550
   ClientWidth     =   8835
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTestSearchBar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   8835
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   2835
      Index           =   5
      Left            =   1200
      ScaleHeight     =   2835
      ScaleWidth      =   3435
      TabIndex        =   49
      Top             =   2640
      Visible         =   0   'False
      Width           =   3435
      Begin VB.CommandButton cmdAction 
         Caption         =   "Info"
         Height          =   375
         Index           =   12
         Left            =   120
         TabIndex        =   82
         Top             =   2400
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   51
         Top             =   600
         Width           =   3255
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use Encryption"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   80
         Top             =   1125
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use Compression"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   77
         Top             =   885
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   54
         Top             =   2040
         Width           =   3255
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "&Browse"
         Height          =   375
         Index           =   4
         Left            =   2100
         TabIndex        =   53
         Top             =   2400
         Width           =   1275
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "&Browse"
         Height          =   375
         Index           =   3
         Left            =   2100
         TabIndex        =   50
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label lbl 
         Height          =   405
         Index           =   13
         Left            =   120
         TabIndex        =   55
         Top             =   1605
         Width           =   3255
      End
      Begin VB.Label lbl 
         Height          =   375
         Index           =   12
         Left            =   120
         TabIndex        =   52
         Top             =   120
         Width           =   3255
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   2220
      Index           =   1
      Left            =   0
      ScaleHeight     =   2220
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   3435
      Begin VB.ComboBox cmbDates 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         ItemData        =   "frmTestSearchBar.frx":09AA
         Left            =   1980
         List            =   "frmTestSearchBar.frx":09B7
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   1890
         Width           =   1080
      End
      Begin MSComCtl2.UpDown udDate 
         Height          =   270
         Left            =   1635
         TabIndex        =   43
         Top             =   1920
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   476
         _Version        =   393216
         Value           =   2
         BuddyControl    =   "txtDate"
         BuddyDispid     =   196612
         OrigLeft        =   1560
         OrigTop         =   1200
         OrigRight       =   1800
         OrigBottom      =   1515
         Max             =   100
         Min             =   2
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.ComboBox cmbDates 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         ItemData        =   "frmTestSearchBar.frx":09D0
         Left            =   120
         List            =   "frmTestSearchBar.frx":09E0
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   1560
         Width           =   1035
      End
      Begin MSComCtl2.DTPicker dtpDates 
         Height          =   315
         Index           =   0
         Left            =   1260
         TabIndex        =   19
         Top             =   1560
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "M/d/yyyy"
         Format          =   19202051
         CurrentDate     =   37982
      End
      Begin VB.ComboBox cmbDates 
         Height          =   315
         Index           =   0
         ItemData        =   "frmTestSearchBar.frx":0A04
         Left            =   1260
         List            =   "frmTestSearchBar.frx":0A11
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optDates 
         Caption         =   "Specify dates"
         Height          =   195
         Index           =   4
         Left            =   0
         TabIndex        =   5
         Top             =   1320
         Width           =   2715
      End
      Begin VB.OptionButton optDates 
         Caption         =   "During the last year"
         Height          =   195
         Index           =   3
         Left            =   0
         TabIndex        =   4
         Top             =   1080
         Width           =   2715
      End
      Begin VB.OptionButton optDates 
         Caption         =   "During the last month"
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   3
         Top             =   840
         Width           =   2715
      End
      Begin VB.OptionButton optDates 
         Caption         =   "During the last week"
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   2
         Top             =   600
         Width           =   2715
      End
      Begin VB.OptionButton optDates 
         Caption         =   "Don't Care"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2715
      End
      Begin VB.TextBox txtDate 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1275
         TabIndex        =   42
         Text            =   "2"
         Top             =   1890
         Visible         =   0   'False
         Width           =   630
      End
      Begin MSComCtl2.DTPicker dtpDates 
         Height          =   315
         Index           =   1
         Left            =   1260
         TabIndex        =   20
         Top             =   1890
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "M/d/yyyy"
         Format          =   19202051
         CurrentDate     =   37982
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Find Files"
         Height          =   255
         Index           =   5
         Left            =   60
         TabIndex        =   40
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   1950
         Visible         =   0   'False
         Width           =   1035
      End
   End
   Begin VB.PictureBox pic 
      Height          =   435
      Index           =   7
      Left            =   5280
      ScaleHeight     =   375
      ScaleWidth      =   2955
      TabIndex        =   68
      Top             =   1920
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   1875
      Index           =   4
      Left            =   3480
      ScaleHeight     =   1875
      ScaleWidth      =   3435
      TabIndex        =   47
      Top             =   360
      Visible         =   0   'False
      Width           =   3435
      Begin VB.CommandButton cmdAction 
         Caption         =   "&Cancel"
         Height          =   495
         Index           =   11
         Left            =   150
         TabIndex        =   81
         Top             =   1320
         Width           =   1395
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "&Start"
         Height          =   495
         Index           =   2
         Left            =   1965
         TabIndex        =   48
         Top             =   1320
         Width           =   1395
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Number of files:"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   76
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label lbl 
         Caption         =   "123,123,456"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   1800
         TabIndex        =   75
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Size:"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   74
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lbl 
         Caption         =   "123,123,456"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   1800
         TabIndex        =   73
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Number of folders:"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   72
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lbl 
         Caption         =   "123,123,456"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   1800
         TabIndex        =   71
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   3120
      Index           =   3
      Left            =   5160
      ScaleHeight     =   3120
      ScaleWidth      =   3435
      TabIndex        =   16
      Top             =   2520
      Visible         =   0   'False
      Width           =   3435
      Begin MSComCtl2.UpDown udChunkSize 
         Height          =   270
         Left            =   2460
         TabIndex        =   85
         Top             =   2835
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   476
         _Version        =   393216
         Value           =   10
         BuddyControl    =   "txtChunkSize"
         BuddyDispid     =   196617
         OrigLeft        =   1560
         OrigTop         =   1200
         OrigRight       =   1800
         OrigBottom      =   1515
         Increment       =   5
         Max             =   100
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   0   'False
      End
      Begin VB.TextBox txtChunkSize 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2100
         TabIndex        =   86
         Text            =   "10"
         Top             =   2805
         Width           =   630
      End
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clear comboboxes without confirmation"
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   78
         Top             =   2520
         Width           =   3255
      End
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Filename implicity enclosed in * wildcards"
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   70
         Top             =   2205
         Width           =   3255
      End
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Auto-Complete"
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   65
         Top             =   1890
         Width           =   2415
      End
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clear list before searching"
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   46
         Top             =   1260
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search within current results"
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   45
         Top             =   1575
         Width           =   2415
      End
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Include Temporary Files"
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   36
         Top             =   945
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Include System Files"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   35
         Top             =   630
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Include Hidden Files"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   34
         Top             =   315
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Include Read-Only Files"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show early results every                files"
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   84
         Top             =   2835
         Width           =   3255
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   2955
      Index           =   6
      Left            =   2880
      ScaleHeight     =   2955
      ScaleWidth      =   3675
      TabIndex        =   57
      Top             =   3000
      Visible         =   0   'False
      Width           =   3675
      Begin VB.CommandButton cmdAction 
         Caption         =   "Clear Selection"
         Enabled         =   0   'False
         Height          =   375
         Index           =   10
         Left            =   1800
         TabIndex        =   79
         Top             =   2520
         Width           =   1635
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "Clear All"
         Enabled         =   0   'False
         Height          =   375
         Index           =   9
         Left            =   0
         TabIndex        =   64
         Top             =   2520
         Width           =   1635
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "Store Selected Files"
         Enabled         =   0   'False
         Height          =   495
         Index           =   8
         Left            =   1800
         TabIndex        =   63
         Top             =   1920
         Width           =   1635
      End
      Begin VB.ListBox lstStored 
         Enabled         =   0   'False
         Height          =   1035
         ItemData        =   "frmTestSearchBar.frx":0A32
         Left            =   0
         List            =   "frmTestSearchBar.frx":0A34
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   62
         ToolTipText     =   "Double-click on an item to restore the results."
         Top             =   840
         Width           =   3435
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "Store All Files"
         Enabled         =   0   'False
         Height          =   495
         Index           =   7
         Left            =   0
         TabIndex        =   61
         Top             =   1920
         Width           =   1635
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "Save To File"
         Enabled         =   0   'False
         Height          =   495
         Index           =   5
         Left            =   0
         TabIndex        =   60
         Top             =   0
         Width           =   1635
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "Load From File"
         Height          =   495
         Index           =   6
         Left            =   1800
         TabIndex        =   59
         Top             =   0
         Width           =   1635
      End
      Begin VB.Label lbl 
         Caption         =   "Temporarily stored lists:"
         Height          =   255
         Index           =   14
         Left            =   0
         TabIndex        =   58
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   1920
      Index           =   2
      Left            =   0
      ScaleHeight     =   1920
      ScaleWidth      =   3555
      TabIndex        =   8
      Top             =   4920
      Visible         =   0   'False
      Width           =   3555
      Begin VB.ComboBox cmbSize 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         ItemData        =   "frmTestSearchBar.frx":0A36
         Left            =   2490
         List            =   "frmTestSearchBar.frx":0A43
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1230
         Width           =   885
      End
      Begin VB.TextBox txtSize 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1200
         TabIndex        =   23
         Text            =   "0"
         Top             =   1590
         Visible         =   0   'False
         Width           =   1005
      End
      Begin MSComCtl2.UpDown udSize 
         Height          =   315
         Index           =   0
         Left            =   2190
         TabIndex        =   22
         Top             =   1230
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtSize(0)"
         BuddyDispid     =   196621
         BuddyIndex      =   0
         OrigLeft        =   1560
         OrigTop         =   1200
         OrigRight       =   1800
         OrigBottom      =   1515
         Max             =   2140000000
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtSize 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   15
         Text            =   "0"
         Top             =   1230
         Width           =   990
      End
      Begin VB.ComboBox cmbSize 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         ItemData        =   "frmTestSearchBar.frx":0A56
         Left            =   120
         List            =   "frmTestSearchBar.frx":0A66
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1230
         Width           =   1035
      End
      Begin VB.OptionButton optSize 
         Caption         =   "Specify size"
         Height          =   195
         Index           =   4
         Left            =   0
         TabIndex        =   13
         Top             =   960
         Width           =   2715
      End
      Begin VB.OptionButton optSize 
         Caption         =   "Large (more than 5 MB)"
         Height          =   195
         Index           =   3
         Left            =   0
         TabIndex        =   12
         Top             =   720
         Width           =   2715
      End
      Begin VB.OptionButton optSize 
         Caption         =   "Medium (between 500 KB and 5 MB)"
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   11
         Top             =   480
         Width           =   3435
      End
      Begin VB.OptionButton optSize 
         Caption         =   "Small (less than 500 KB)"
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   10
         Top             =   240
         Width           =   2715
      End
      Begin VB.OptionButton optSize 
         Caption         =   "Don't Care"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   2715
      End
      Begin MSComCtl2.UpDown udSize 
         Height          =   315
         Index           =   1
         Left            =   2190
         TabIndex        =   24
         Top             =   1590
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtSize(1)"
         BuddyDispid     =   196621
         BuddyIndex      =   1
         OrigLeft        =   1560
         OrigTop         =   1200
         OrigRight       =   1800
         OrigBottom      =   1515
         Max             =   2140000000
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   21
         Top             =   1650
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin ComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   56
      Top             =   6735
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
      EndProperty
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   2355
      Index           =   0
      Left            =   0
      ScaleHeight     =   2355
      ScaleWidth      =   3495
      TabIndex        =   27
      Top             =   240
      Visible         =   0   'False
      Width           =   3495
      Begin VB.ComboBox cmbFile 
         Height          =   315
         Index           =   2
         Left            =   0
         TabIndex        =   32
         ToolTipText     =   "Text search is not case sensitive, but correct spelling is required."
         Top             =   2040
         Width           =   3435
      End
      Begin VB.ComboBox cmbFile 
         Height          =   315
         Index           =   1
         Left            =   0
         TabIndex        =   31
         ToolTipText     =   "Use * for multi-letter and ? for single-letter wildcards."
         Top             =   1440
         Width           =   3435
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "&Browse"
         Height          =   375
         Index           =   1
         Left            =   2160
         TabIndex        =   30
         ToolTipText     =   "Click to Browse for the folder to be searched."
         Top             =   720
         Width           =   1275
      End
      Begin VB.ComboBox cmbFile 
         Height          =   315
         Index           =   0
         Left            =   0
         TabIndex        =   29
         Top             =   360
         Width           =   3435
      End
      Begin VB.CheckBox chkRecurse 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Include Sub-Folders"
         Height          =   255
         Left            =   0
         TabIndex        =   28
         ToolTipText     =   "Expands the search to return files from sub-folders, sub-subfolders, etc."
         Top             =   750
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.Label lbl 
         Caption         =   "Choose a folder in which to begin your search."
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   39
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label lbl 
         Caption         =   "A word or phrase in the file."
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   38
         Top             =   1800
         Width           =   3375
      End
      Begin VB.Label lbl 
         Caption         =   "Part or all of the file name."
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   37
         Top             =   1200
         Width           =   3375
      End
   End
   Begin vbalExplorerBarLib6.vbalExplorerBarCtl barSearch 
      Align           =   3  'Align Left
      Height          =   6735
      Left            =   0
      TabIndex        =   67
      Top             =   0
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   11880
      BackColorEnd    =   0
      BackColorStart  =   0
   End
   Begin ComctlLib.ListView lv 
      Height          =   2145
      Left            =   5880
      TabIndex        =   41
      Top             =   960
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   3784
      View            =   3
      SortOrder       =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      OLEDropMode     =   1
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin MSComCtl2.Animation ani 
      Height          =   1095
      Left            =   4920
      TabIndex        =   83
      Top             =   2400
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1931
      _Version        =   393216
      Center          =   -1  'True
      BackColor       =   16777215
      FullWidth       =   249
      FullHeight      =   73
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   1035
      Index           =   8
      Left            =   5040
      ScaleHeight     =   1035
      ScaleWidth      =   3555
      TabIndex        =   17
      Top             =   5640
      Width           =   3555
      Begin VB.CommandButton cmdAction 
         Caption         =   "Search"
         Default         =   -1  'True
         Height          =   855
         Index           =   0
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   90
         Width           =   2355
      End
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "Preparing File Information..."
      Height          =   255
      Index           =   16
      Left            =   5400
      TabIndex        =   69
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   15
      Left            =   5160
      TabIndex        =   66
      Top             =   60
      Width           =   2775
   End
   Begin VB.Menu mnuFile 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuAction 
         Caption         =   "List Items"
         Index           =   0
         Begin VB.Menu mnuListItem 
            Caption         =   "Refresh"
            Index           =   0
         End
         Begin VB.Menu mnuListItem 
            Caption         =   "Select All"
            Index           =   1
         End
         Begin VB.Menu mnuListItem 
            Caption         =   "Invert Selection"
            Index           =   2
         End
         Begin VB.Menu mnuListItem 
            Caption         =   "Remove Selection"
            Index           =   3
         End
         Begin VB.Menu mnuListItem 
            Caption         =   "Clear"
            Index           =   4
         End
      End
      Begin VB.Menu mnuAction 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuAction 
         Caption         =   "All Shown Files"
         Index           =   2
         Begin VB.Menu mnuAll 
            Caption         =   "Show Total Size"
            Index           =   0
         End
         Begin VB.Menu mnuAll 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnuAll 
            Caption         =   "Backup to Composite File"
            Index           =   2
         End
         Begin VB.Menu mnuAll 
            Caption         =   "Move to Folder"
            Index           =   3
         End
         Begin VB.Menu mnuAll 
            Caption         =   "Copy to Folder"
            Index           =   4
         End
         Begin VB.Menu mnuAll 
            Caption         =   "Delete"
            Index           =   5
         End
      End
      Begin VB.Menu mnuAction 
         Caption         =   "Selected Files"
         Index           =   3
         Begin VB.Menu mnuSel 
            Caption         =   "Show Total Size"
            Index           =   0
         End
         Begin VB.Menu mnuSel 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnuSel 
            Caption         =   "Backup to Composite File"
            Index           =   2
         End
         Begin VB.Menu mnuSel 
            Caption         =   "Restore Composite File(s)"
            Index           =   3
         End
         Begin VB.Menu mnuSel 
            Caption         =   "Move to Folder"
            Index           =   4
         End
         Begin VB.Menu mnuSel 
            Caption         =   "Copy to Folder"
            Index           =   5
         End
         Begin VB.Menu mnuSel 
            Caption         =   "Delete"
            Index           =   6
         End
      End
      Begin VB.Menu mnuAction 
         Caption         =   "Selected File"
         Index           =   4
         Begin VB.Menu mnuOne 
            Caption         =   "Open"
            Index           =   0
         End
         Begin VB.Menu mnuOne 
            Caption         =   "Open Containing Folder"
            Index           =   1
         End
         Begin VB.Menu mnuOne 
            Caption         =   "Edit with Notepad"
            Index           =   2
         End
         Begin VB.Menu mnuOne 
            Caption         =   "Properties"
            Index           =   3
         End
         Begin VB.Menu mnuOne 
            Caption         =   "Rename"
            Index           =   4
         End
         Begin VB.Menu mnuOne 
            Caption         =   "Print"
            Index           =   5
         End
      End
      Begin VB.Menu mnuAction 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuAction 
         Caption         =   "Load File List..."
         Index           =   6
      End
      Begin VB.Menu mnuAction 
         Caption         =   "Save File List..."
         Index           =   7
      End
      Begin VB.Menu mnuAction 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuAction 
         Caption         =   "View"
         Index           =   9
         Begin VB.Menu mnuView 
            Caption         =   "Large Icons"
            Index           =   0
         End
         Begin VB.Menu mnuView 
            Caption         =   "Small Icons"
            Index           =   1
         End
         Begin VB.Menu mnuView 
            Caption         =   "List"
            Index           =   2
         End
         Begin VB.Menu mnuView 
            Caption         =   "Details"
            Index           =   3
         End
         Begin VB.Menu mnuView 
            Caption         =   "Columns..."
            Index           =   4
         End
      End
      Begin VB.Menu mnuAction 
         Caption         =   "Arrange"
         Index           =   10
         Begin VB.Menu mnuArrange 
            Caption         =   "Top"
            Index           =   0
         End
         Begin VB.Menu mnuArrange 
            Caption         =   "Left"
            Index           =   1
         End
         Begin VB.Menu mnuArrange 
            Caption         =   "None"
            Index           =   2
         End
      End
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements iFileTaskParent
Implements iRichDialogParent

Private Enum eMessages
    msgLoadFile
    msgGetTempSlotName
    msgLoadTempSlot
    msgClearList
    msgBlankStartFolder
    msgFileError
    msgNameAlreadyUsed
    msgConfirmDelete
    msgRelativePathInfo
End Enum

Private Enum ePages
    pgSearch
    pgBackup
    pgRestore
    pgCopy
    pgMove
End Enum

Private Enum eMnuView
    mnuLarge
    mnuSmall
    mnuList
    mnuDetails
    mnuColumns
End Enum

Private Enum eMnuArrange
    mnuTop
    mnuLeft
    mnuNone
End Enum

Private Enum eMnuAction
    mnuListItems
    
    mnuAllShown = mnuListItems + 2
    mnuSelected
    mnuSelectedOne
    
    mnuLoadSearch = mnuSelectedOne + 2
    mnuSaveSearch
    
    mnuView = mnuSaveSearch + 2
    mnuArrange
End Enum

Private Enum eMnuListItems
    mnuRefresh
    mnuliSelectAll
    mnuliInvertSelection
    mnuliRemoveSelection
    mnuClear
End Enum

Private Enum eSelectedFilesActions
    mnuShowTotals
    mnuBackup = mnuShowTotals + 2
    mnuRestore
    mnuMove
    mnuCopy
    mnuDelete
End Enum

Private Enum eSelectedFileActions
    mnuOpen
    mnuOpenContaining
    mnuEdit
    mnuProps
    mnuRename
    mnuPrint
End Enum

Private Enum eLbl
    lblDate
    lblSize
    lblSearchFolder
    lblSearchFile
    lblSearchText
    lblDateType
    lblTaskNumFiles
    lblTaskNumFolders
    lblTaskTotalSize
    lblTaskNumFiles2
    lblTaskNumFolders2
    lblTaskTotalSize2
    lblDestinationPath
    lblRelativePath
    lblTempStorage
    lblLVBack
    lblPleaseWait
End Enum

Private Enum eCmbFiles
    cmbFilePath
    cmbFileName
    cmbFileText
End Enum

Private Enum eOptSize
    optSizeAny
    optSizeSmall
    optSizeMedium
    optSizeLarge
    optSizeCustom
End Enum

Private Enum eOptDates
    optDatesAny
    optDatesWeek
    optDatesMonth
    optDatesYear
    optDatesCustom
End Enum

Private Enum eCmd
    cmdStart
    cmdBrowseSearch
    cmdTaskStart
    cmdBrowseDestination
    cmdBrowseRelative
    cmdSave
    cmdLoad
    cmdTempStoreAll
    cmdTempStoreSel
    cmdClearTempAll
    cmdClearTempSel
    cmdTaskCancel
    cmdRelativePathInfo
End Enum

Private Enum eControlPair
    cpOne
    cpTwo
End Enum

Private Enum ePic
    picFiles
    picDates
    picSize
    picOptions
    picTaskHeader
    picTaskPaths
    picStorage
    picProgress
    picSearch
End Enum

Private Enum eRangeSelectionType
    rsAtLeast
    rsAtMost
    rsBetween
    rsNear
End Enum

Private Enum eFileScales
    fsByte
    fsKilo
    fsMega
End Enum

Private Enum eChkOptions
    chkReadOnly
    chkHidden
    chkSystem
    chkTemp
    chkClearList
    chkSubSearch
    chkAutoComplete
    chkEncloseWildcard
    chkSilentClearCombo
    chkChunksize
End Enum

Private Enum eUIStates
    uiNormal
    uiSearching
    uiWorking
    uiModalInput
End Enum

Private Const LVM_FIRST = &H1000
Private Const LVM_GETITEMCOUNT = (LVM_FIRST + 4)
Private Const LVM_GETSELECTEDCOUNT = (LVM_FIRST + 50)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private WithEvents moLVUtils As cLVUtils
Attribute moLVUtils.VB_VarHelpID = -1

Private moFiles As cFiles
Private moTaskFiles As cFiles
Private moProgress As cProgressBar
Private moFindFiles As cFileSearch
Private moFileLV As cFileListView
Private moTempStorage As Collection

Private miKeyDown(cmbFilePath To cmbFileText) As Byte

Private miFilesFound As Long
Private mbSyncFlag As Boolean
Private mbAutoComplete As Boolean
Private miPage As ePages
Private mbCancel As Boolean
Private mbSilentClearCombo As Boolean
Private mbSearching As Boolean
Private mbSelectedOnly As Boolean
Private mbLoaded As Boolean
Private miUIState As eUIStates

Private Sub InitSearchBar()
    Dim liBackcolor As OLE_COLOR
    Dim liTemp As OLE_COLOR
    liBackcolor = vbBlue
    With barSearch
        With .Bars
            .Clear
            barSearch.Redraw = False
            Select Case miPage
                Case pgSearch
                    UIState = uiNormal
                    AddPicToBar picFiles
                    AddPicToBar picDates
                    AddPicToBar picSize
                    AddPicToBar picOptions
                    AddPicToBar picStorage
        
                    liTemp = pic(picFiles).BackColor
                    chkRecurse.BackColor = liTemp

                    With lbl
                        .Item(lblSearchFolder).BackColor = liTemp
                        .Item(lblSearchFile).BackColor = liTemp
                        .Item(lblSearchText).BackColor = liTemp
                        
                        SetBackcolor optDates, liTemp
                        .Item(lblDate).BackColor = liTemp
                        .Item(lblDateType).BackColor = liTemp
                        SetBackcolor optSize, liTemp
                        .Item(lblSize).BackColor = liTemp
                    
                        SetBackcolor chkOptions, liTemp
                        
                        .Item(lblTempStorage).BackColor = liTemp
                    End With
            
                Case Else
                    UIState = uiModalInput
                    
                    AddPicToBar picTaskHeader
                    AddPicToBar picTaskPaths
                    
                    Dim lsDest As String
                    Dim lsRelative As String
                    
                    Select Case miPage
                        Case pgRestore
                            lsDest = "Enter the directory into which the contents of each composite file will be extracted."
                            lsRelative = ""
                        Case pgMove
                            lsDest = "Enter the directory into which the files will be moved."
                            lsRelative = ""
                        Case pgCopy
                            lsDest = "Enter the directory into which the files will be copied."
                            lsRelative = ""
                        Case pgBackup
                            lsDest = "Enter the new composite file that will be created from the files."
                            lsRelative = "Enter the relative path to be compared to the each of the file's paths."
                    End Select
                    
                    liTemp = pic(picTaskHeader).BackColor

                    With moTaskFiles
                        lbl.Item(lblTaskNumFiles).Caption = .Count
                        lbl.Item(lblTaskNumFolders).Caption = .GetFolders.Count
                        lbl.Item(lblTaskTotalSize).Caption = Format(.TotalSize \ KB, "###,###,###,###,##0 KB")
                        txt(cpOne).Text = ""
                        txt(cpTwo).Text = .CommonRoot
                    End With
                    Dim lbVal As Boolean
                    lbVal = Len(lsRelative) > 0
                    
                    With lbl
                        SetBackcolor chk, liTemp
                        .Item(lblTaskNumFiles).BackColor = liTemp
                        .Item(lblTaskNumFolders).BackColor = liTemp
                        .Item(lblTaskTotalSize).BackColor = liTemp
                        .Item(lblTaskNumFiles2).BackColor = liTemp
                        .Item(lblTaskNumFolders2).BackColor = liTemp
                        .Item(lblTaskTotalSize2).BackColor = liTemp
                        With .Item(lblDestinationPath)
                            .Caption = lsDest
                            .BackColor = liTemp
                        End With
                        With .Item(lblRelativePath)
                            .Caption = lsRelative
                            .BackColor = liTemp
                            .Visible = lbVal
                        End With
                        txt(cpTwo).Visible = lbVal
                        cmdAction(cmdBrowseRelative).Visible = lbVal
                        cmdAction(cmdRelativePathInfo).Visible = lbVal
                        lbVal = miPage = pgBackup Or miPage = pgRestore
                        chk(cpOne).Visible = lbVal
                        chk(cpTwo).Visible = lbVal
                        
                    End With
                End Select
        End With
        .Redraw = True
    End With
End Sub

Private Sub AddPicToBar(piPic As ePic)
    Dim liBackcolor As Long
    Dim lbCanExpand As Boolean
    Dim lbIsSpecial As Boolean
    Dim liTitleForeColor As Long
    Dim liTitleForeColorOver As Long
    Dim lsKey As String
    Dim lsCaption As String
    Dim liState As EExplorerBarStates
    Dim loPic As PictureBox
    
    Set loPic = pic(piPic)
    
    liBackcolor = vbBlue
    Select Case piPic
        Case picTaskPaths
            liState = eBarExpanded
            liTitleForeColor = vbWhite
            liTitleForeColorOver = vbWhite
            lbCanExpand = False
            lsKey = "PATH"
            lsCaption = "Enter a Path"
        Case picTaskHeader
            liState = eBarExpanded
            liTitleForeColor = vbWhite
            liTitleForeColorOver = vbWhite
            lbCanExpand = False
            lbIsSpecial = True
            lsKey = "TASK"
            lsCaption = "Task Information"
        Case picStorage
            liState = eBarCollapsed
            liTitleForeColor = vbWhite
            liTitleForeColorOver = vbWhite
            lbCanExpand = True
            lsKey = "STORAGE"
            lsCaption = "Save/Load File Lists"
        Case picSize
            liState = eBarCollapsed
            liTitleForeColor = vbWhite
            liTitleForeColorOver = vbWhite
            lbCanExpand = True
            lsKey = "SIZE"
            lsCaption = "File Size"
        Case picOptions
            liState = eBarCollapsed
            liTitleForeColor = vbWhite
            liTitleForeColorOver = vbWhite
            lbCanExpand = True
            lsKey = "OPTIONS"
            lsCaption = "Other Search Options"
        Case picFiles
            liState = eBarExpanded
            liTitleForeColor = vbBlack
            liTitleForeColorOver = vbBlack
            lbIsSpecial = True
            lbCanExpand = False
            lsKey = "FILES"
            lsCaption = "Search by any or all of these criteria."
        Case picDates
            liState = eBarCollapsed
            liTitleForeColor = vbWhite
            liTitleForeColorOver = vbWhite
            lbCanExpand = True
            lsKey = "PICDATES"
            lsCaption = "File Date/Time"
    End Select
    With barSearch.Bars.Add(, lsKey, lsCaption)
        .State = liState
        .BackColor = liBackcolor
        .CanExpand = lbCanExpand
        .IsSpecial = lbIsSpecial
        .TitleForeColor = liTitleForeColor
        .TitleForeColorOver = liTitleForeColorOver
        .Items.Add(, lsKey, , , eItemControlPlaceHolder).Control = loPic
    End With
End Sub

Private Sub SetBackcolor(poControls As Object, piColor As OLE_COLOR)
    On Error Resume Next
    Dim loControl
    For Each loControl In poControls
        loControl.BackColor = piColor
    Next
End Sub

Private Sub SetEnabled(poControls As Object, pbVal As Boolean)
    On Error Resume Next
    Dim loControl
    For Each loControl In poControls
        loControl.Enabled = pbVal
    Next
End Sub

Private Sub SetBold(poControls As Object)
    On Error Resume Next
    Dim loControl As OptionButton
    For Each loControl In poControls
        loControl.Font.Bold = loControl.Value = True
    Next
End Sub

Private Sub chkOptions_Click(Index As Integer)
    Select Case Index
        Case chkSubSearch
            cmbFile(cmbFilePath).Enabled = chkOptions(Index).Value = 0
            cmdAction(cmdBrowseSearch).Enabled = chkOptions(Index).Value = 0
        Case chkAutoComplete
            mbAutoComplete = chkOptions(Index).Value = 1
        Case chkSilentClearCombo
            mbSilentClearCombo = chkOptions(Index).Value = 1
        Case chkChunksize
            txtChunkSize.Enabled = chkOptions(Index).Value = 1
            udChunkSize.Enabled = chkOptions(Index).Value = 1
    End Select
End Sub

Private Sub cmbDates_Click(Index As Integer)
    Dim liIndex As eRangeSelectionType
    Dim lbVal As Boolean
    
    liIndex = cmbDates(Index).ListIndex
    
    Select Case Index
        Case cpTwo
            If liIndex = rsNear Then lbVal = True
            txtDate.Visible = lbVal
            udDate.Visible = lbVal
            cmbDates(2).Visible = lbVal
            NormalRangeSelection liIndex, lbl(lblDate), dtpDates(cpTwo)
            If lbVal Then dtpDates(cpTwo).Visible = False
    End Select
End Sub

Private Function SetRangeSelection(piRange As eFindFileRange)
    On Error Resume Next
    Dim ldblLow As Double
    Dim ldblHigh As Double
    Dim ldblDiff As Double
    Dim ldblTemp As Double
   
    Dim liType As eRangeSelectionType
    liType = -1
    
    Select Case piRange
        Case ffrSize
            If optSize(optSizeAny).Value Then
                ldblLow = 0
                ldblHigh = 0
            ElseIf optSize(optSizeSmall).Value Then
                ldblLow = 0
                ldblHigh = 511999
            ElseIf optSize(optSizeMedium).Value Then
                ldblLow = 512000
                ldblHigh = 5242879
            ElseIf optSize(optSizeLarge).Value Then
                ldblLow = 5242880
                ldblHigh = 0
            ElseIf optSize(optSizeCustom).Value Then
                ldblLow = ScaleVal(Val(txtSize(cpOne).Text), cmbSize(cpTwo).ListIndex, fsByte)
                ldblHigh = ScaleVal(Val(txtSize(cpTwo).Text), cmbSize(cpTwo).ListIndex, fsByte)
                liType = cmbSize(cpOne).ListIndex
            End If
        Case Else
            If optDates(optDatesAny).Value Then
                ldblLow = 0
                ldblHigh = 0
            ElseIf optDates(optDatesWeek).Value Then
                ldblLow = DateAdd("ww", -1, Now)
                ldblHigh = Now
            ElseIf optDates(optDatesMonth).Value Then
                ldblLow = DateAdd("m", -1, Now)
                ldblHigh = Now
            ElseIf optDates(optDatesYear).Value Then
                ldblLow = DateAdd("m", -12, Now)
                ldblHigh = Now
            Else
                ldblLow = CDbl(dtpDates(cpOne).Value)
                ldblHigh = CDbl(dtpDates(cpTwo).Value)
                liType = cmbDates(cpTwo).ListIndex
            End If
    End Select
    

    Select Case liType
        Case rsNear
            If piRange = ffrSize Then
                ldblDiff = ldblHigh
                ldblTemp = ldblLow
                ldblLow = ldblTemp - ldblDiff
                ldblHigh = ldblTemp + ldblDiff
            Else
                ldblDiff = Val(txtDate.Text)
                ldblHigh = ldblLow
                Select Case cmbDates(2).ListIndex
                    Case 0 'Days
                        ldblLow = ldblHigh - ldblDiff
                        ldblHigh = ldblHigh + ldblDiff
                    Case 1 'Weeks
                        ldblLow = DateAdd("ww", -ldblDiff, ldblHigh)
                        ldblHigh = DateAdd("ww", ldblDiff, ldblHigh)
                    Case 2 'Months
                        ldblLow = DateAdd("m", -ldblDiff, ldblHigh)
                        ldblHigh = DateAdd("m", ldblDiff, ldblHigh)
                End Select
            End If
        Case rsAtMost
            ldblHigh = ldblLow
            ldblLow = 0
        Case rsAtLeast
            ldblHigh = 0
    End Select
    
    If ldblLow > ldblHigh And ldblHigh <> 0 Then
        ldblDiff = ldblHigh
        ldblHigh = ldblLow
        ldblLow = ldblDiff
    End If
    
    moFindFiles.SetRange piRange, ldblLow, ldblHigh
End Function

Private Sub NormalRangeSelection(ByVal Index As eRangeSelectionType, ByVal poLabel As Label, ByVal poControl1 As Control, Optional ByVal poControl2 As Control)
    Dim lbVal As Boolean
    Dim lsVal As String
    On Error Resume Next
    Select Case Index
        Case rsAtLeast
            lbVal = False
        Case rsAtMost
            lbVal = False
        Case rsBetween
            lbVal = True
            lsVal = "And:"
        Case rsNear
            lbVal = True
            lsVal = "Variance:"
    End Select
    poLabel.Visible = lbVal
    poLabel.Caption = lsVal
    poControl1.Visible = lbVal
    poControl2.Visible = lbVal
End Sub

Private Sub cmbFile_Change(Index As Integer)
    On Error Resume Next
    With cmbFile(Index)
        Select Case Index
            Case cmbFilePath
                If mbSyncFlag Or Not mbAutoComplete Then Exit Sub
                mbSyncFlag = True
                ac_Change cmbFile(cmbFilePath), miKeyDown(Index), acbFolder Or acbMultiSelect
                mbSyncFlag = False
            Case Else
                If mbSyncFlag Or Not mbAutoComplete Then Exit Sub
                mbSyncFlag = True
                ac_Change cmbFile(Index), miKeyDown(Index)
                mbSyncFlag = False
        End Select
        .Tag = .Text
    End With
End Sub

Private Sub cmbFile_Click(Index As Integer)
    If mbSyncFlag Then Exit Sub
    Dim lsTemp As String
    With cmbFile(Index)
        If .ListCount = .ListIndex + 1 Then
            mbSyncFlag = True
            If .ListCount > 1 Then
                .Text = .Tag
                .SelStart = Len(.Text)
                If ShowMessage(msgClearList) = rdYes Then
                    .Clear
                    .AddItem "Clear This List"
                End If
                lsTemp = .List(0)
                .List(0) = .Text
                .ListIndex = 0
                .List(0) = lsTemp
            End If
            .Text = .Tag
            .SelStart = Len(.Text)
            mbSyncFlag = False
            .Tag = .Text
            TrashNextSetText GetFirstChild(.hwnd)
        Else
            .Tag = .Text
        End If
    End With
End Sub

Private Sub cmbFile_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    miKeyDown(Index) = KeyCode
    PathKeyDown
End Sub

Private Sub cmbFile_KeyPress(Index As Integer, KeyAscii As Integer)
    ac_KeyPress cmbFile(Index), KeyAscii, IIf(Index = cmbFilePath, acbFolder + acbMultiSelect, acbListOnly + acbMultiSelect)
End Sub

Private Sub cmbSize_Click(Index As Integer)
    On Error Resume Next
    Dim liIndex As Long
    liIndex = cmbSize(Index).ListIndex
    Select Case Index
        Case cpOne
            NormalRangeSelection liIndex, lbl(lblSize), txtSize(cpTwo), udSize(cpTwo)
        Case cpTwo
            udSize(cpOne).Value = ScaleVal(txtSize(cpOne).Text, txtSize(cpOne).Tag, liIndex)
            udSize(cpTwo).Value = ScaleVal(txtSize(cpTwo).Text, txtSize(cpOne).Tag, liIndex)
            txtSize(cpOne).Tag = liIndex
            liIndex = 2 ^ (Abs(liIndex - 2) * 8)
            udSize(cpOne).Increment = liIndex
            udSize(cpTwo).Increment = liIndex
    End Select
End Sub

Private Function ScaleVal(Val, ByVal piFrom As eFileScales, ByVal piTo As eFileScales)
    On Error Resume Next
    Dim liPower As Long
    liPower = piTo - piFrom
    ScaleVal = Val
    Select Case liPower
        Case Is > 0
            For liPower = 1 To liPower
                ScaleVal = ScaleVal \ KB
            Next
        Case Is < 0
            For liPower = 1 To Abs(liPower)
                ScaleVal = ScaleVal * KB
            Next
    End Select
End Function

Private Sub cmdAction_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case cmdStart
            Dim loTask As iFileTask
            Set loTask = moFindFiles
            If cmdAction(cmdStart).Caption = "&Stop" Then loTask.Canceled = True: mbCancel = True Else StartSearch
        Case cmdBrowseSearch
            BrowseFoldertoTextBox cmbFile(cmbFilePath), "Choose a folder to begin your search."
        Case cmdBrowseDestination
            If miPage = pgBackup Then
                BrowseFiletoTextBox txt(cpOne), "Browse For File", CommonDialogFilter("CryptoZip Composite Files (*.czf)", "*.czf", "All Files (*)", "*"), ".czf"
            Else
                BrowseFoldertoTextBox txt(cpOne), "Choose a destination folder."
            End If
        Case cmdBrowseRelative
            BrowseFoldertoTextBox txt(cpTwo), "Choose a folder to serve as the relative starting point."
        Case cmdTaskStart, cmdTaskCancel
            If Index = cmdTaskStart Then
                Dim lsDest As String, lsRelative As String
                Dim lbVal As Boolean
                Dim loUI As frmFileTask
                Dim liTask As Long
                Dim lbEncrypt As Boolean
                Dim lbCompress As Boolean
                'Set loUI = New frmFileTask
                lsDest = txt(cpOne).Text
                If txt(cpTwo).Visible Then lsRelative = txt(cpTwo).Text
                lbCompress = chk(cpOne).Value = 1
                lbEncrypt = chk(cpTwo).Value = 1
                Select Case miPage
                    Case pgRestore
                        If lbCompress And lbEncrypt Then
                            liTask = ftCryptoUnZip
                        ElseIf lbCompress Then
                            liTask = ftUnzip
                        ElseIf lbEncrypt Then
                            liTask = ftDecrypt
                        End If
                    Case pgBackup
                        If lbCompress And lbEncrypt Then
                            liTask = ftCryptoZip
                        ElseIf lbCompress Then
                            liTask = ftZip
                        ElseIf lbEncrypt Then
                            liTask = ftEncrypt
                        End If
                    Case pgMove
                        liTask = ftMove
                    Case pgCopy
                        liTask = ftCopy
                End Select
                If ShowTask(moTaskFiles, liTask, lsDest, lsRelative) Then
                    If miPage = pgMove Then RemoveItems
                End If
            End If
            miPage = pgSearch
            InitSearchBar
        Case cmdTempStoreAll
            SaveTempSlot GetFiles(False)
        Case cmdTempStoreSel
            SaveTempSlot GetFiles(True)
        Case cmdSave
            SaveResults
        Case cmdLoad
            LoadResults
        Case cmdClearTempSel
            Dim i As Long
            Dim lsTemp As String
            For i = lstStored.ListCount - 1 To 0 Step -1
                If lstStored.Selected(i) Then
                    lsTemp = lstStored.List(i)
                    lstStored.RemoveItem i
                    moTempStorage.Remove lsTemp & gsKeySuffix
                End If
            Next
            UpdateLSTStats
        Case cmdClearTempAll
            lstStored.Clear
            Set moTempStorage = New Collection
            UpdateLSTStats
        Case cmdRelativePathInfo
            ShowMessage msgRelativePathInfo
    End Select
End Sub

Private Sub RemoveItems()
    On Error Resume Next
    If mbSelectedOnly Then
        Dim loLI As ListItem
        Dim loColl As Collection
        Dim lvTemp
        
        Set loColl = New Collection
        
        For Each loLI In lv.ListItems
            If loLI.Selected Then loColl.Add loLI.Key
        Next
        With lv.ListItems
            For Each lvTemp In loColl
                .Remove lvTemp
                moFiles.RemoveFiles lvTemp
            Next
        End With
    Else
        lv.ListItems.Clear
        moFiles.Clear
    End If
    UpdateLVStats
End Sub

Private Sub BrowseFiletoTextBox(poText As Object, psTitle As String, psFilter As String, psExt As String)
    On Error Resume Next
    Dim lsString As String
    lsString = poText.Text
    If Not FolderExists(lsString) Then lsString = PathGetParentFolder(lsString)
    lsString = GetSaveFileName(hwnd, psTitle, poText.Text, psFilter, psExt, OFN_EXPLORER + OFN_HIDEREADONLY)
    If Len(lsString) > 0 Then poText.Text = lsString
End Sub

Private Sub BrowseFoldertoTextBox(poText As Object, psTitle As String)
    On Error Resume Next
    mbSyncFlag = True
    Dim lsTemp As String
    With poText
        lsTemp = BrowseForFolder(hwnd, psTitle, .Text)
        If Len(lsTemp) > 0 Then .Text = lsTemp
    End With
    mbSyncFlag = False
End Sub

Private Sub StartSearch()
    If Len(cmbFile(cmbFilePath).Text) = 0 Then
        ShowMessage msgBlankStartFolder
        cmbFile(cmbFilePath).SetFocus
        Exit Sub
    End If
    mbCancel = False
    Dim lvTemp, i As Long, lsTemp As String
    With moFindFiles
        For i = ffiReadOnly To ffiTemporary
            .Ignore(i) = chkOptions(i).Value = 0
        Next
        lsTemp = cmbFile(cmbFileName).Text
        If chkOptions(chkEncloseWildcard).Value = 1 Then
            If Right$(lsTemp, 1) <> "*" Then lsTemp = lsTemp & "*"
            If Left$(lsTemp, 1) <> "*" Then lsTemp = "*" & lsTemp
        End If
        .Filter = lsTemp
        
        .Path = cmbFile(cmbFilePath).Text
        .FindContainedText = cmbFile(cmbFileText).Text
        .Recursive = chkRecurse.Value = 1
        lvTemp = cmbDates(cpOne).ListIndex
        For i = ffrAccessed To ffrModified
            If i = lvTemp Then
                SetRangeSelection i
            Else
                moFindFiles.SetRange i, 0, 0
            End If
        Next
        SetRangeSelection ffrSize
        For i = cmbFilePath To cmbFileText
            AddExclusiveItem cmbFile(i)
        Next
        'test
        i = 0
        If chkOptions(chkChunksize).Value = 1 Then i = udChunkSize.Value
        .ChunkSize = i
        If chkOptions(chkSubSearch).Value = 1 Then Set .SearchNames = GetFileNames
        Dim loTask As iFileTask
        Set loTask = moFindFiles
        If chkOptions(chkClearList).Value = 1 Then lv.ListItems.Clear
        miFilesFound = 0
        loTask.Start
    End With
End Sub

Private Function GetFileNames(Optional pbSelected As Boolean) As Collection
    Dim loItem As ListItem
    Dim lsTemp As String
    Set GetFileNames = New Collection
    
    With GetFileNames
        If pbSelected Then
            For Each loItem In lv.ListItems
                If loItem.Selected Then
                    lsTemp = loItem.Key
                    .Add lsTemp, lsTemp
                End If
            Next
        Else
            For Each loItem In lv.ListItems
                lsTemp = loItem.Key
                .Add lsTemp, lsTemp
            Next
        End If
    End With
End Function

Private Function GetFiles(Optional pbSelected As Boolean) As cFiles
    Dim loItem As ListItem
    Dim lsTemp As String
    If pbSelected Then
        Set GetFiles = moFiles.Clone(GetFileNames(pbSelected))
    Else
        Set GetFiles = moFiles.Clone
    End If
End Function

Private Sub Form_Initialize()
    On Error Resume Next
    Set moTempStorage = New Collection
    Set moLVUtils = New cLVUtils
    moLVUtils.Attach lv
    LoadColHeaders
    LoadFile
    Set moProgress = New cProgressBar
    Set moFiles = New cFiles
    With moProgress
        .DrawObject = pic(picProgress)
        .ShowText = True
        .BackColor = vbButtonFace
        .BarColor = vbHighlight
    End With
    Set moFileLV = New cFileListView
    moFileLV.Attach lv
    'moFileLV.Columns = flvName Or flvSize Or flvType Or flvFolder Or flvModified
    Set moFindFiles = CreateObject(cSearchClass)
    Dim loTask As iFileTask
    Set loTask = moFindFiles
    Set loTask.Parent = Me
    'barSearch.BackColorEnd = RGB(172, 172, 172)
    'barSearch.BackColorStart = RGB(172, 172, 172)
    'barSearch.UseExplorerStyle = True
    InitSearchBar
    With sb.Panels
        .Add(, "last", "", sbrText).AutoSize = sbrSpring
        .Add(, "total", "", sbrText).AutoSize = sbrContents
        .Add(, "selected", "", sbrText).AutoSize = sbrContents
    End With
    
    With cmbFile
        SetPathbreakProc GetFirstChild(.Item(cmbFilePath).hwnd)
        SetPathbreakProc GetFirstChild(.Item(cmbFileName).hwnd)
        SetPathbreakProc GetFirstChild(.Item(cmbFileText).hwnd)
    End With
    
    With txt
        SetPathbreakProc .Item(cpOne).hwnd
        SetPathbreakProc .Item(cpTwo).hwnd
    End With
    
    cmdAction(cmdStart).Picture = LoadResPicture(102, vbResIcon)
    
    UpdateLVStats
End Sub

Private Sub LoadResults()
    Dim loBag As cFiles
    Dim loColl As Collection
    Dim lsFile As String
    
    'Const liTempSlot = -101&
    On Error Resume Next
    Set loColl = New Collection
    If GetOpenFileNames(loColl, hwnd, "Load Results From File", "", CommonDialogFilter("Find Results Files(*.frf)", "*.frf", "All Files (*.*)", "*.*"), "czb", OFN_EXPLORER + OFN_FILEMUSTEXIST + OFN_HIDEREADONLY) Then
        lsFile = loColl(1)
        Set loBag = New cFiles
        If loBag.FileSaveLoad(lsFile, False) Then LoadBag loBag, ShowMessage(msgLoadFile) Else ShowMessage msgFileError
    End If
End Sub

Private Sub LoadBag(ByVal poBag As cFiles, piAnswer As eRichDialogReturn)
    On Error Resume Next
    Dim liType As eFileCombineLogic
    Const liTempSlot = -101&
    Select Case piAnswer
        Case rdButton1
            liType = fclUnion
        Case rdButton2
            liType = fclIntersection
        Case rdButton3
            liType = fclExclusion
        Case rdButton4
            liType = fclNegation
        Case rdButton5
            liType = liTempSlot
        Case Else
            Exit Sub
    End Select
    
    If liType <> liTempSlot Then
        UIState = uiWorking
        Dim i As Long
        Dim liCount As Long
        Dim loEachFile As cFile
        
        Set poBag = GetFiles(False).Combine(poBag, liType)
        lv.ListItems.Clear
        moFiles.Clear
        liCount = poBag.Count
        For Each loEachFile In poBag
            i = i + 1
            With loEachFile
                moFiles.AddObject loEachFile
                moFileLV.ShowFile .FullPath, .Attributes, .Size, .Modified, .Accessed, .Created
                UpdateProgress i, liCount
            End With
            If Not mbLoaded Then Exit Sub
            If mbCancel Then Exit For
        Next
        UIState = uiNormal
    Else
        SaveTempSlot poBag
    End If
    
End Sub

Private Sub SaveResults()
    Dim lsFile As String
     
    lsFile = GetSaveFileName(hwnd, "Save Results To File", "", CommonDialogFilter("Find Results Files (*.frf)", "*.frf", "All Files (*.*)", "*.*"), ".frf", OFN_EXPLORER + OFN_OVERWRITEPROMPT + OFN_NOREADONLYRETURN + OFN_HIDEREADONLY)
    If Len(lsFile) > 0 Then
        If Not GetFiles(False).FileSaveLoad(lsFile, True) Then ShowMessage msgFileError
    End If
End Sub

Private Sub SaveTempSlot(ByVal poBag As cFiles)
    Dim lsName As String
    On Error Resume Next
again:
    lsName = InputBoxEx("Enter the name that you would like to identify these files with.", rdOKCancel + rdCancelRaiseError + rdCancelButton2 + rdDefaultButton1 + rdDisallowBlankInput, "Choose a Temporary Name", , , hwnd, rdCenterCenter)
    If Err.Number = 0 Then
        If Len(lsName) > 0 Then
            moTempStorage.Add poBag, lsName & gsKeySuffix
            If Err.Number = 0 Then
                lstStored.AddItem lsName
            Else
                If ShowMessage(msgNameAlreadyUsed) = rdYes Then GoTo again
            End If
        End If
    End If
    UpdateLSTStats
End Sub

Private Sub LoadTempSlot()
    On Error Resume Next
    Dim lsTemp As String
    lsTemp = lstStored.List(lstStored.ListIndex)
    If Len(lsTemp) > 0 Then
        Dim loBag As cFiles
        Set loBag = moTempStorage.Item(lsTemp & gsKeySuffix)
        If Not loBag Is Nothing Then LoadBag loBag, ShowMessage(msgLoadTempSlot)
    End If
End Sub

Private Function ShowMessage(piMessage As eMessages) As eRichDialogReturn
    Dim lsMsg As String
    Dim lsTitle As String
    Dim liAtt As eRichDialogAttributes
    Dim lvButtons
    Dim liTag As eMessages
    Select Case piMessage
        Case msgClearList
            If mbSilentClearCombo Then
                ShowMessage = rdYes
                Exit Function
            End If
            lsMsg = "Do you really want to clear all the items from this list?"
            liAtt = rdQuestion + rdYesNo Or rdDefaultButton2 Or rdCancelButton2
            lsTitle = "Confirm Clear List"
            liTag = msgClearList
        Case msgLoadTempSlot
            lsMsg = Replace(LoadResString(102), "%TempName%", lstStored.List(lstStored.ListIndex), 1, 1)
            lvButtons = Array("Union", "Intersection", "Exclusion", "Negation", "&Cancel")
            liAtt = rdQuestion Or rdDefaultButton4 Or rdCancelButton4
            lsTitle = "Combine Results"
        Case msgLoadFile
            lsMsg = LoadResString(101)
            liAtt = rdQuestion Or rdDefaultButton5 Or rdCancelButton6
            lvButtons = Array("Union", "Intersection", "Exclusion", "Negation", "Temp Slot", "&Cancel")
            lsTitle = "Load From File"
        Case msgBlankStartFolder
            lsMsg = "You must choose a folder to begin the search."
            liAtt = rdInformation Or rdDefaultButton1
            lsTitle = "Where to begin?"
        Case msgConfirmDelete
            lsMsg = ListCount(mbSelectedOnly)
            lsMsg = "Are you sure that you want to delete these " & lsMsg & " items?"
            liAtt = rdCritical + rdYesNo + rdDefaultButton2 + rdCancelButton2
            lsTitle = "Confirm Delete"
        Case msgRelativePathInfo
            lsMsg = "When you create a composite file (any file that contains data from multiple files) that includes files from multiple folders, you may wish to preserve the folder tree upon extraction.  To help with this, part of each file's current path can be saved with the it; the relative path to the folder you choose.  All of the paths are saved as relative to that one folder so that when they are extracted to a different folder or on a different machine they keep the sub-folder structure intact.  If the relative path is left blank, then the sub-folder tree is disregarded and all files will be placed into the same destination folder when they are extracted."
            lsTitle = "Relative Path Information"
    End Select
    ShowMessage = MsgBoxEx(lsMsg, liAtt, lsTitle, , Me.hwnd, rdCenterCenter, lvButtons, , liTag, Me)
End Function

Private Function AccessFile(pbLoad As Boolean) As Boolean
    Dim loFile As cFileIO
    Dim lsFileName As String
    Dim lyTemp() As Byte
    Dim i As Long
    Dim j As Long
    
    On Error Resume Next
    'lsFileName = GetSetting("VBFind", "Ini", "Filename", "")
    lsFileName = PathBuild(App.Path, "VBFind.ini")
    'If Len(lsFileName) = 0 Then
        'lsFileName = PathBuild(PathGetSpecial(sfTemporary), "VBFind.dat")
        'SaveSetting "VBFind", "Ini", "Filename", lsFileName
    'End If
    
    Set loFile = New cFileIO

    With loFile
        If pbLoad Then
            .FileAccess = GENERIC_READ
            .FileCreation = OPEN_EXISTING
            .FileFlags = FILE_FLAG_SEQUENTIAL_SCAN
            .FileShare = FILE_SHARE_READ
            If .OpenFile(lsFileName) Then
                AccessFile = True
                LoadControls cmbFile, ctComboText, loFile
                LoadControls optDates, ctOption, loFile
                LoadControls cmbDates, ctCombo, loFile
                LoadControls dtpDates, ctDateTime, loFile
                LoadControls optSize, ctOption, loFile
                LoadControls cmbSize, ctCombo, loFile
                LoadControls udSize, ctUpDown, loFile
                LoadControls chkOptions, ctCheck, loFile
                txtSize(cpOne).Text = udSize(cpOne).Value
                txtSize(cpTwo).Text = udSize(cpTwo).Value
                loFile.GetLong i
                udChunkSize.Value = i
                txtChunkSize.Text = i
                loFile.GetLong i
                loFile.GetLong j
                loFile.GetBytes lyTemp, j
                'DecompressByteArray lyTemp, i
                moLVUtils.ColumnHeaders.SetColData lyTemp
            End If
            moLVUtils.ColumnHeaders.Redraw = True
            moFileLV.SyncColumns
            With cmbFile
                For i = .LBound To .UBound
                    .Item(i).AddItem "Clear This List"
                Next
            End With
            chkOptions(chkSubSearch).Value = 0
        Else
            .FileAccess = GENERIC_WRITE
            .FileCreation = CREATE_ALWAYS
            .FileFlags = FILE_FLAG_SEQUENTIAL_SCAN
            .FileShare = FILE_SHARE_READ
            If .OpenFile(lsFileName) Then
                AccessFile = True
                With cmbFile
                    For i = .LBound To .UBound
                        .Item(i).RemoveItem .Item(i).ListCount - 1
                    Next
                End With
                
                udSize(cpOne).Value = CLng(txtSize(cpOne).Text)
                udSize(cpTwo).Value = CLng(txtSize(cpTwo).Text)
                SaveControls cmbFile, ctComboText, loFile
                SaveControls optDates, ctOption, loFile
                SaveControls cmbDates, ctCombo, loFile
                SaveControls dtpDates, ctDateTime, loFile
                SaveControls optSize, ctOption, loFile
                SaveControls cmbSize, ctCombo, loFile
                SaveControls udSize, ctUpDown, loFile
                SaveControls chkOptions, ctCheck, loFile
                loFile.AppendLong udChunkSize.Value
                With moLVUtils.ColumnHeaders
                    .GetColData lyTemp
                    loFile.AppendLong UBound(lyTemp) + 1
                    'CompressByteArray lyTemp
                    loFile.AppendLong UBound(lyTemp) + 1
                    loFile.AppendBytes lyTemp
                End With
            End If
        End If
    
    End With
End Function

Private Sub LoadColHeaders()
    With moLVUtils.ColumnHeaders
        .Redraw = False
        .Clear
        .Add "Name", "Name", 1800, True
        .Add "In Folder", "In Folder", 2000, True
        .Add "Type", "Type", 1600, True
        .Add "Modified", "Modified", 1800, True
        .Add "Accessed", "Accessed", 1800, False
        .Add "Created", "Created", 1800, False
        .Add "Size", "Size", 600, True
        .Add "Attributes", "Attributes", 800, False
    End With
End Sub

Private Sub LoadFile()
    If Not AccessFile(True) Then
        optDates(optDatesAny).Value = True
        cmbDates(cpOne).ListIndex = 2
        cmbDates(cpTwo).ListIndex = rsAtLeast
        cmbDates(2).ListIndex = 2
        dtpDates(cpOne).Value = DateAdd("m", -6, Date)
        dtpDates(cpTwo).Value = Date
        optSize(optSizeAny).Value = True
        cmbSize(cpOne).ListIndex = rsAtLeast
        cmbSize(cpTwo).ListIndex = fsKilo
        udSize(cpOne).Value = 1024
        udSize(cpTwo).Value = 128
        chkOptions(chkAutoComplete).Value = 1
    End If
End Sub

Private Sub Form_Load()
    mbLoaded = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    
    lv.Move barSearch.Width, 0, Width - barSearch.Width - 100, Height - 420 - pic(picSearch).Height - sb.Height
    With lv
        lbl(lblLVBack).Move .Left, .Top, .Width, .Height
        lbl(lblPleaseWait).Move .Left + .Width \ 2 - lbl(lblPleaseWait).Width \ 2, .Height \ 2 + .Top - lbl(lblPleaseWait).Height
        lbl(lblLVBack).Refresh
        ani.Move .Left + .Width \ 2 - ani.Width \ 2, .Height \ 2 + .Top - ani.Height \ 2
        pic(picProgress).Width = lv.Width * 0.6
        pic(picProgress).Move .Left + .Width \ 2 - pic(picProgress).Width \ 2, .Top + .Height \ 2
    End With
    
    With pic(picSearch)
        .Move barSearch.Width + (Width - barSearch.Width) \ 2 - .Width \ 2, lv.Height, .Width, .Height
        .Refresh
    End With
    
End Sub

Private Sub Form_Terminate()
    Set moProgress = Nothing
    Set moFileLV = Nothing
    Set moFindFiles = Nothing
    Set moLVUtils = Nothing
    Set moTempStorage = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbLoaded = False
    mbCancel = True
    Unload frmLVColumns
    AccessFile False
End Sub

Private Sub iFileTaskParent_Notify(Tag As Long)
    On Error Resume Next
    Dim loTask As iFileTask
    Dim loFile As cFile
    
    Set loTask = moFindFiles
    Select Case loTask.Status Mod ftsCanceled
        Case ftsJustStarting
            UIState = uiSearching
        Case ftsFinishing
            If moFindFiles.ChunkSize = 0 Then
                Dim liTotal As Long
                Dim liCount As Long
                liTotal = loTask.Files.Count
                UIState = uiWorking
                For Each loFile In loTask.Files
                    moFiles.AddObject loFile
                    With loFile
                        moFileLV.ShowFile .FullPath, .Attributes, .Size, .Modified, .Accessed, .Created
                        liCount = liCount + 1
                        UpdateProgress liCount, liTotal
                    End With
                    If Not mbLoaded Then Exit Sub
                    If mbCancel Then Exit For
                Next
                miFilesFound = liTotal
            Else
                For Each loFile In loTask.Files
                    moFiles.AddObject loFile
                    With loFile
                        moFileLV.ShowFile .FullPath, .Attributes, .Size, .Modified, .Accessed, .Created
                    End With
                Next
                miFilesFound = miFilesFound + loTask.Files.Count
            End If
            UIState = uiNormal
            loTask.Files.Clear
            UpdateLVStats
        Case Else
            For Each loFile In loTask.Files
                moFiles.AddObject loFile
                With loFile
                    moFileLV.ShowFile .FullPath, .Attributes, .Size, .Modified, .Accessed, .Created
                End With
            Next
            miFilesFound = miFilesFound + loTask.Files.Count
            loTask.Files.Clear
            UpdateLVStats
    End Select
End Sub

Private Sub iRichDialogParent_HasReturned(ByVal Dialog As RichDialogs.iRichDialog)
    Dim liMsg As eMessages
    liMsg = Dialog.Info.Tag
    If liMsg = msgClearList Then chkOptions(chkSilentClearCombo).Value = Abs(Dialog.Info.CheckBoxValue)
End Sub

Private Sub iRichDialogParent_QueryInfo(ByVal Dialog As RichDialogs.iRichDialog, bCancel As Boolean)
    Dim liMsg As eMessages
    liMsg = Dialog.Info.Tag
    If liMsg = msgClearList Then
        With Dialog.Info
            .CheckBoxValue = mbSilentClearCombo
            .CheckBoxStatement = "Do not show this message again"
        End With
    End If
End Sub

Private Sub iRichDialogParent_WillShow(ByVal Dialog As RichDialogs.iRichDialog, bCancel As Boolean)
'
End Sub

Private Sub lstStored_DblClick()
    LoadTempSlot
End Sub

Private Sub lv_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error Resume Next
    Dim lsName As String
    Dim lsNew As String
    Dim liIndex As Long
    With lv.SelectedItem
        lsName = .Key
        liIndex = .Index
    End With
    lsNew = PathBuild(PathGetParentFolder(lsName), NewString)
    If FileMove(lsName, lsNew) Then
        moFiles.RemoveFiles lsName
        moFiles.AddFile lsNew
        With lv.ListItems
            .Remove liIndex
            .Add liIndex, lsNew, PathGetFileName(lsNew)
        End With
        With moFiles(lsNew)
            moFileLV.ShowFile lsNew, .Attributes, .Size, .Modified, .Accessed, .Created
        End With
        Set lv.SelectedItem = lv.ListItems(liIndex)
    Else
        Cancel = True
    End If
End Sub

Private Sub lv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 93 Then lv_MouseUp 2, 0, 0, 0
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UpdateLVStats
    If Button = 2 Then
        Dim lbSel As Boolean
        Dim lbAny As Boolean
        If Not mbSearching Then
            lbAny = ListCount(False) > 0
            If lbAny Then lbSel = ListCount(True) > 0
        
            With mnuAction
                .Item(mnuListItems).Enabled = lbAny And Not mbSearching
                mnuListItem(mnuliRemoveSelection).Enabled = lbSel
            
                .Item(mnuAllShown).Enabled = lbAny
                .Item(mnuSelected).Enabled = lbSel
                .Item(mnuSelectedOne).Enabled = Not lv.SelectedItem Is Nothing
            
                .Item(mnuLoadSearch).Enabled = Not mbSearching
                .Item(mnuSaveSearch).Enabled = lbAny
                .Item(mnuView).Enabled = Not mbSearching
                .Item(mnuArrange).Enabled = Not mbSearching
            End With
            PopupMenu mnuFile, vbPopupMenuRightButton ', , , mnuAction(mnuOpen)
        End If
    End If
End Sub

Private Sub UpdateLVStats()
    On Error Resume Next
    Dim liAll As Long
    Dim liSel As Long
    Dim lbVal As Boolean
    With sb.Panels
        liAll = SendMessage(lv.hwnd, LVM_GETITEMCOUNT, 0&, 0&)
        liSel = SendMessage(lv.hwnd, LVM_GETSELECTEDCOUNT, 0&, 0&)
        .Item("total").Text = "Total Files: " & liAll
        .Item("last").Text = "Files found in last search: " & miFilesFound
        .Item("selected").Text = "Files Selected: " & liSel
    End With
    lbVal = liAll > 0
    chkOptions(chkSubSearch).Enabled = lbVal
    cmdAction(cmdSave).Enabled = lbVal
    cmdAction.Item(cmdTempStoreAll).Enabled = lbVal
    cmdAction.Item(cmdTempStoreSel).Enabled = liSel > 0
End Sub

Private Sub UpdateLSTStats()
    Dim lbVal As Boolean
    With cmdAction
        lbVal = lstStored.ListCount > 0
        .Item(cmdSave).Enabled = ListCount(False) > 0
        .Item(cmdClearTempSel).Enabled = lstStored.SelCount > 0
        '.Item(cmdClearTempAll).Enabled = lstStored.SelCount > 0
        .Item(cmdClearTempAll).Enabled = lbVal
        lstStored.Enabled = lbVal
    End With
End Sub

Private Sub lv_OLECompleteDrag(Effect As Long)
    Dim lvTemp
    Dim loColl As Collection
    On Error Resume Next
    Set loColl = GetFileNames(True)
    For Each lvTemp In loColl
        If Not FileExists(CStr(lvTemp)) Then
            lv.ListItems.Remove lvTemp
            moFiles.RemoveFiles lvTemp
        End If
    Next
    lv.Tag = ""
    Effect = 0
End Sub

Private Sub lv_OLEDragDrop(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    Dim loTemp As Collection
    Dim ltFind As tFindFiles
    Dim lvTemp
    On Error Resume Next
    With Data
        For i = 1 To .Files.Count
            If FolderExists(.Files(i)) Then
                ltFind.Path = .Files(i)
                With ltFind
                    .Filter = "*"
                    .Recurse = True
                End With
                Set loTemp = FindFiles(ltFind)
                For Each lvTemp In loTemp
                    moFiles.AddFile lvTemp
                    With moFiles(lvTemp)
                        moFileLV.ShowFile CStr(lvTemp), .Attributes, .Size, .Modified, .Accessed, .Created
                    End With
                Next
            Else
                moFiles.AddFile .Files(i)
                With moFiles(.Files(i))
                    moFileLV.ShowFile .FullPath, .Attributes, .Size, .Modified, .Accessed, .Created
                End With
            End If
        Next
    End With
End Sub

Private Sub lv_OLEDragOver(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    On Error Resume Next
    If LenB(lv.Tag) = 0 Then
        Effect = vbDropEffectCopy
    Else
        Effect = 0
    End If
End Sub

Private Sub lv_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    On Error Resume Next
    If LenB(lv.Tag) = 0 Then
        Effect = vbDropEffectCopy
        DefaultCursors = True
    Else
        Effect = 0
    End If
End Sub
    
Private Sub lv_OLEStartDrag(Data As ComctlLib.DataObject, AllowedEffects As Long)
    On Error Resume Next
    lv.Tag = " "
    AllowedEffects = vbDropEffectCopy Or vbDropEffectMove
    Dim loColl As Collection, lvTemp
    Set loColl = GetFileNames(True)
    With Data
        For Each lvTemp In loColl
            .Files.Add lvTemp
        Next
        .SetData , vbCFFiles
    End With
End Sub

Private Sub mnuAction_Click(Index As Integer)
    Select Case Index
        Case mnuSaveSearch
            SaveResults
        Case mnuLoadSearch
            LoadResults
    End Select
End Sub

Private Sub mnuAll_Click(Index As Integer)
    On Error Resume Next
    mbSelectedOnly = False
    Dim liIndex As Long
    liIndex = Index
    If liIndex >= mnuRestore Then liIndex = liIndex + 1
    Select Case liIndex
        Case mnuShowTotals
            ShowTotals
        Case mnuBackup
            miPage = pgBackup
        Case mnuMove
            miPage = pgMove
        Case mnuCopy
            miPage = pgCopy
        Case mnuDelete
            If ShowMessage(msgConfirmDelete) = rdYes Then
                If ShowTask(GetFiles(False), ftDelete, "", "") Then RemoveItems
            End If
            Exit Sub
    End Select
    InitSearchBar
End Sub

Private Sub mnuArrange_Click(Index As Integer)
    Select Case Index
        Case mnuTop
            lv.Arrange = lvwAutoTop
        Case mnuLeft
            lv.Arrange = lvwAutoLeft
        Case mnuNone
            lv.Arrange = lvwNone
    End Select
End Sub

Private Sub mnuListItem_Click(Index As Integer)
    On Error Resume Next
    Dim loLI As ListItem
    Dim loColl As Collection
    Dim lvTemp As Variant
    Dim liCount As Long
    Dim liTotal As Long
    Select Case Index
        Case mnuliSelectAll
            For Each loLI In lv.ListItems
                loLI.Selected = True
            Next
            UpdateLVStats
        Case mnuliInvertSelection
            For Each loLI In lv.ListItems
                loLI.Selected = Not loLI.Selected
            Next
            UpdateLVStats
        Case mnuliRemoveSelection
            Set loColl = New Collection
            For Each loLI In lv.ListItems
                If loLI.Selected Then loColl.Add loLI.Key
            Next
            With lv.ListItems
                For Each lvTemp In loColl
                    .Remove lvTemp
                    moFiles.RemoveFiles lvTemp
                Next
            End With
            UpdateLVStats
        Case mnuClear
            lv.ListItems.Clear
            moFiles.Clear
        Case mnuRefresh
            UIState = uiWorking
            liTotal = lv.ListItems.Count
            Set loColl = New Collection
            moFiles.Clear
            For Each loLI In lv.ListItems
                liCount = liCount + 1
                UpdateProgress liCount, liTotal
                If FileExists(loLI.Key) Then
                    moFiles.AddFile loLI.Key
                    With moFiles(loLI.Key)
                        moFileLV.ShowFile .FullPath, .Attributes, .Size, .Modified, .Accessed, .Created
                    End With
                Else
                    loColl.Add loLI.Key
                End If
                If Not mbLoaded Then Exit Sub
                If mbCancel Then Exit For
            Next
            With lv.ListItems
                For Each lvTemp In loColl
                    .Remove lvTemp
                Next
            End With
            UIState = uiNormal
    End Select
End Sub

Private Sub mnuOne_Click(Index As Integer)
    Dim lsName As String
    On Error Resume Next
    lsName = lv.SelectedItem.Key
    If Err.Number <> 0 Then Exit Sub
    Select Case Index
        Case mnuOpen
            shelldoc lsName, stOpen
        Case mnuOpenContaining
            shelldoc PathGetParentFolder(lsName), stExplore
        Case mnuEdit
            shellexe "notepad.exe", lsName
        Case mnuProps
            ShowProps lsName
        Case mnuRename
            lv.StartLabelEdit
        Case mnuPrint
            shelldoc lsName, stPrint
    End Select
End Sub

Private Sub ShowTotals()
    On Error Resume Next
    Dim ldblTotal As Double
    Const NumFormat = "###,###,###,###,###,##0"
    With GetFiles(mbSelectedOnly)
        ldblTotal = .TotalSize
        MsgBoxEx IIf(mbSelectedOnly, "You have selected ", "There are ") & .Count & " files, totaling " & Format$(ldblTotal, NumFormat) & " bytes (" & Format(ldblTotal / KB, NumFormat & ".00") & " kb) in " & .GetFolders.Count & " folders.", , "File Size"
    End With
End Sub

Private Sub mnuSel_Click(Index As Integer)
    On Error Resume Next
    mbSelectedOnly = True
    Select Case Index
        Case mnuShowTotals
            ShowTotals
        Case mnuBackup
            miPage = pgBackup
        Case mnuRestore
            miPage = pgRestore
        Case mnuMove
            miPage = pgMove
        Case mnuCopy
            miPage = pgCopy
        Case mnuDelete
            If ShowMessage(msgConfirmDelete) = rdYes Then
                If ShowTask(GetFiles(True), ftDelete, "", "") Then RemoveItems
            End If
            Exit Sub
    End Select
    InitSearchBar
End Sub

Private Sub mnuView_Click(Index As Integer)
    Select Case Index
        Case mnuLarge
            lv.View = lvwIcon
        Case mnuSmall
            lv.View = lvwSmallIcon
        Case mnuList
            lv.View = lvwList
        Case mnuDetails
            lv.View = lvwReport
        Case mnuColumns
            frmLVColumns.ChooseColumns moLVUtils.ColumnHeaders, Me
            moFileLV.SyncColumns
    End Select
End Sub

Private Sub moLVUtils_Drag(ByVal Buttons As Long, ByVal Shift As Long)
    UpdateLVStats
End Sub

Private Sub moLVUtils_ItemActivated()
    mnuOne_Click mnuOpen
End Sub

Private Sub moLVUtils_ItemSelected(ByVal Item As ComctlLib.ListItem)
    UpdateLVStats
End Sub

Private Sub optDates_Click(Index As Integer)
    Dim lbVal As Boolean
    SetBold optDates
    lbVal = Not optDates(optDatesAny).Value
    With cmbDates(cpOne)
        lbl(lblDateType).FontBold = lbVal
        cmbDates(cpOne).FontBold = lbVal
        .Enabled = lbVal
        .BackColor = IIf(.Enabled, vbWindowBackground, vbButtonFace)
    End With
    
    lbVal = optDates(optDatesCustom).Value
    SetEnabled dtpDates, lbVal
    With cmbDates(cpTwo)
        .Enabled = lbVal
        .BackColor = IIf(lbVal, vbWindowBackground, vbButtonFace)
    End With
    With cmbDates(2)
        .Enabled = lbVal
        .BackColor = IIf(lbVal, vbWindowBackground, vbButtonFace)
    End With
    With txtDate
        .Enabled = lbVal
        .BackColor = IIf(lbVal, vbWindowBackground, vbButtonFace)
    End With
End Sub

Private Sub optSize_Click(Index As Integer)
    Dim lbVal As Boolean
    lbVal = optSize(optSizeCustom).Value
    
    SetBold optSize
    SetEnabled cmbSize, lbVal
    SetBackcolor cmbSize, IIf(lbVal, vbWindowBackground, vbButtonFace)
    SetEnabled txtSize, lbVal
    SetBackcolor txtSize, IIf(lbVal, vbWindowBackground, vbButtonFace)
    SetEnabled udSize, lbVal
    
End Sub

Private Sub pic_Paint(Index As Integer)
    If Index = picProgress Then
        On Error Resume Next
        moProgress.Draw
    End If
End Sub

Private Sub txt_Change(Index As Integer)
    If mbSyncFlag Or Not mbAutoComplete Then Exit Sub
    mbSyncFlag = True
    ac_Change txt(Index), Val(txt(Index).Tag), IIf(Index = cpOne And miPage = pgBackup, acbFile, acbFolder)
    mbSyncFlag = False
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    txt(Index).Tag = KeyCode
    PathKeyDown
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    ac_KeyPress txt(Index), KeyAscii, IIf(Index = 1 And miPage = pgBackup, acbFile, acbFolder)
End Sub

Private Sub txtChunkSize_Change()
    On Error Resume Next
    udChunkSize.Value = Val(txtChunkSize.Text)
End Sub

Private Sub txtChunkSize_KeyPress(KeyAscii As Integer)
    If KeyAscii = 22 Then
        KeyAscii = 0
        PasteOnlyNums txtChunkSize
    Else
        FilterNumericKeyAscii KeyAscii
    End If
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 22 Then
        KeyAscii = 0
        PasteOnlyNums txtDate
    Else
        FilterNumericKeyAscii KeyAscii
    End If
End Sub

Private Sub txtSize_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 22 Then
        KeyAscii = 0
        PasteOnlyNums txtSize(Index)
    Else
        FilterNumericKeyAscii KeyAscii
    End If
End Sub

Public Property Get ListCount(Optional pbSelectedOnly As Boolean) As Long
    If pbSelectedOnly Then
        ListCount = SendMessage(lv.hwnd, LVM_GETSELECTEDCOUNT, 0&, 0&)
    Else
        ListCount = SendMessage(lv.hwnd, LVM_GETITEMCOUNT, 0&, 0&)
    End If
End Property

Private Sub UpdateProgress(ByVal piDone As Long, ByVal piTotal As Long)
    With moProgress
        .Value = piDone
        .Max = piTotal
        .Text = Int(.Percent) & " %"
    End With
    IfUserInputOrPaintThenDoEvents
    'DoEvents
End Sub

Private Property Let UIState(ByVal piVal As eUIStates)
    Dim lbStopButton As Boolean
    Dim lbButtonEnabled As Boolean
    Dim lbListEnabled As Boolean
    Dim lbListVisible As Boolean
    Dim lbAniVisible As Boolean
    Dim lbProgressVisible As Boolean
    Dim lbStorageEnabled As Boolean
    
    On Error Resume Next
    mbCancel = False
    mbSearching = False
    Set moTaskFiles = Nothing
    Select Case piVal
        Case uiWorking
            lbStopButton = True
            lbButtonEnabled = True
            lbListEnabled = True
            lbProgressVisible = True
        Case uiSearching
            lbStopButton = True
            lbButtonEnabled = True
            lbListEnabled = True
            lbListVisible = chkOptions(chkChunksize).Value = 1
            mbSearching = True
            lbAniVisible = Not lbListVisible
        Case uiModalInput
            Set moTaskFiles = GetFiles(mbSelectedOnly)
            lbListVisible = True
            With cmdAction
                .Item(cmdTaskStart).Default = True
                .Item(cmdTaskCancel).Cancel = True
            End With
        Case Else
            lbButtonEnabled = True
            lbListEnabled = True
            lbListVisible = True
            lbStorageEnabled = True
    End Select
    cmdAction(cmdStart).Enabled = lbButtonEnabled
    ShowStartStopButton lbStopButton
    lv.Enabled = lbListEnabled
    lv.Visible = lbListVisible
    pic(picStorage).Enabled = lbStorageEnabled
    If lbAniVisible Then
        ani.Open PathBuild(App.Path, "search.avi")
        ani.Visible = True
        ani.Play
    Else
        ani.Close
        ani.Visible = False
    End If
    pic(picProgress).Visible = lbProgressVisible
    UpdateLVStats
End Property

Private Sub ShowStartStopButton(Optional ByVal pbStop As Boolean)
    Dim lsCaption As String
    Dim loPicture As StdPicture
    Dim lbDef As Boolean
    Dim lbCancel As Boolean
    If pbStop Then
        lsCaption = "&Stop"
        Set loPicture = LoadResPicture(101, vbResIcon)
        lbCancel = True
    Else
        lsCaption = "&Start"
        Set loPicture = LoadResPicture(102, vbResIcon)
        lbDef = True
    End If
    With cmdAction(cmdStart)
        .Caption = lsCaption
        Set .Picture = loPicture
        If .Enabled Then
            .Default = lbDef
            .Cancel = lbCancel
        End If
    End With
End Sub

'Public Sub test()
'    Exit Sub
'    Dim ldbl1 As Double
'    Dim ldbl2 As Double
'    With moFindFiles
'        .GetRange ffrAccessed, ldbl1, ldbl2
'        Debug.Print "Accessed:" & CDate(ldbl1), CDate(ldbl2)
'        .GetRange ffrCreated, ldbl1, ldbl2
'        Debug.Print "Created:" & CDate(ldbl1), CDate(ldbl2)
'        .GetRange ffrModified, ldbl1, ldbl2
'        Debug.Print "Modified:" & CDate(ldbl1), CDate(ldbl2)
'        .GetRange ffrSize, ldbl1, ldbl2
'        Debug.Print "Size:" & ldbl1, ldbl2
'        Debug.Print "Filter:" & .Filter
'        Debug.Print "Text:" & .FindContainedText
'        Debug.Print "IR/O" & .Ignore(ffiReadOnly)
'        Debug.Print "IHidden" & .Ignore(ffiHidden)
'        Debug.Print "ITemp" & .Ignore(ffiTemporary)
'        Debug.Print "ISystem" & .Ignore(ffiSystem)
'        Debug.Print "Path: " & .Path
'        Debug.Print "Recursive: " & .Recursive
'        'Stop
'    End With
'End Sub
