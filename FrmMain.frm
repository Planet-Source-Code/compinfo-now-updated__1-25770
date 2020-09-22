VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6795
   ClientLeft      =   810
   ClientTop       =   825
   ClientWidth     =   7680
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   7680
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   12303
      _Version        =   393216
      Style           =   1
      Tabs            =   14
      TabsPerRow      =   7
      TabHeight       =   617
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Windows"
      TabPicture(0)   =   "FrmMain.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Lbl5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Lbl6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Lbl7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Lbl8"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Lbl2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Lbl3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Lbl4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Lbl1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblA"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblB"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblC"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblD"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lbl9"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Timer1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "CPU && BIOS"
      TabPicture(1)   =   "FrmMain.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl1E"
      Tab(1).Control(1)=   "lbl1D"
      Tab(1).Control(2)=   "lbl1C"
      Tab(1).Control(3)=   "lbl1B"
      Tab(1).Control(4)=   "lbl1A"
      Tab(1).Control(5)=   "lbl19"
      Tab(1).Control(6)=   "lbl18"
      Tab(1).Control(7)=   "lbl17"
      Tab(1).Control(8)=   "Lbl16"
      Tab(1).Control(9)=   "Lbl10"
      Tab(1).Control(10)=   "Lbl11"
      Tab(1).Control(11)=   "Lbl12"
      Tab(1).Control(12)=   "Lbl15"
      Tab(1).Control(13)=   "Lbl14"
      Tab(1).Control(14)=   "Lbl13"
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "Memory"
      TabPicture(2)   =   "FrmMain.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "PB1"
      Tab(2).Control(1)=   "PB2"
      Tab(2).Control(2)=   "PB3"
      Tab(2).Control(3)=   "Lbl26"
      Tab(2).Control(4)=   "Lbl25"
      Tab(2).Control(5)=   "Lbl24"
      Tab(2).Control(6)=   "Lbl23"
      Tab(2).Control(7)=   "Lbl22"
      Tab(2).Control(8)=   "Lbl21"
      Tab(2).Control(9)=   "Lbl20"
      Tab(2).Control(10)=   "Lbl28"
      Tab(2).Control(11)=   "Lbl27"
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "Power"
      TabPicture(3)   =   "FrmMain.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Lbl34"
      Tab(3).Control(1)=   "Lbl33"
      Tab(3).Control(2)=   "Lbl32"
      Tab(3).Control(3)=   "Lbl31"
      Tab(3).Control(4)=   "Lbl30"
      Tab(3).Control(5)=   "lbl35"
      Tab(3).Control(6)=   "lbl36"
      Tab(3).Control(7)=   "lbl37"
      Tab(3).Control(8)=   "lbl38"
      Tab(3).Control(9)=   "lbl39"
      Tab(3).Control(10)=   "lbl3A"
      Tab(3).Control(11)=   "lbl3B"
      Tab(3).ControlCount=   12
      TabCaption(4)   =   "Keyboard"
      TabPicture(4)   =   "FrmMain.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "List6"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Video Info"
      TabPicture(5)   =   "FrmMain.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "List3"
      Tab(5).Control(1)=   "List2"
      Tab(5).Control(2)=   "lbl57"
      Tab(5).Control(3)=   "lbl56"
      Tab(5).Control(4)=   "lbl55"
      Tab(5).Control(5)=   "lbl54"
      Tab(5).Control(6)=   "lbl53"
      Tab(5).Control(7)=   "lbl52"
      Tab(5).Control(8)=   "lbl50"
      Tab(5).Control(9)=   "lbl51"
      Tab(5).ControlCount=   10
      TabCaption(6)   =   " Fonts "
      TabPicture(6)   =   "FrmMain.frx":04EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "lbl62"
      Tab(6).Control(1)=   "lbl61"
      Tab(6).Control(2)=   "lbl60"
      Tab(6).Control(3)=   "List1"
      Tab(6).ControlCount=   4
      TabCaption(7)   =   "Sound Info"
      TabPicture(7)   =   "FrmMain.frx":0506
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "lbl70"
      Tab(7).Control(1)=   "List4"
      Tab(7).ControlCount=   2
      TabCaption(8)   =   "Windows II"
      TabPicture(8)   =   "FrmMain.frx":0522
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "lbl88"
      Tab(8).Control(1)=   "lbl80"
      Tab(8).Control(2)=   "lbl81"
      Tab(8).Control(3)=   "lbl82"
      Tab(8).Control(4)=   "lbl83"
      Tab(8).Control(5)=   "lbl84"
      Tab(8).Control(6)=   "lbl85"
      Tab(8).Control(7)=   "lbl86"
      Tab(8).Control(8)=   "lbl87"
      Tab(8).Control(9)=   "lbl89"
      Tab(8).Control(10)=   "lbl8A"
      Tab(8).Control(11)=   "lbl8B"
      Tab(8).Control(12)=   "lbl8C"
      Tab(8).Control(13)=   "lbl8D"
      Tab(8).Control(14)=   "lbl8E"
      Tab(8).Control(15)=   "lbl8F"
      Tab(8).ControlCount=   16
      TabCaption(9)   =   "Mouse"
      TabPicture(9)   =   "FrmMain.frx":053E
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "List5"
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "Drives"
      TabPicture(10)  =   "FrmMain.frx":055A
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Drive1"
      Tab(10).Control(1)=   "lbl10C"
      Tab(10).Control(2)=   "lbl105"
      Tab(10).Control(3)=   "lbl106"
      Tab(10).Control(4)=   "lbl104"
      Tab(10).Control(5)=   "lbl103"
      Tab(10).Control(6)=   "lbl102"
      Tab(10).Control(7)=   "lbl101"
      Tab(10).Control(8)=   "lbl100"
      Tab(10).Control(9)=   "lbl107"
      Tab(10).Control(10)=   "lbl108"
      Tab(10).Control(11)=   "lbl109"
      Tab(10).Control(12)=   "lbl10A"
      Tab(10).Control(13)=   "lbl10B"
      Tab(10).ControlCount=   14
      TabCaption(11)  =   "Dirs"
      TabPicture(11)  =   "FrmMain.frx":0576
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "lbl110"
      Tab(11).Control(1)=   "lbl111"
      Tab(11).Control(2)=   "lbl112"
      Tab(11).Control(3)=   "lbl113"
      Tab(11).Control(4)=   "lbl114"
      Tab(11).Control(5)=   "lbl117"
      Tab(11).Control(6)=   "lbl116"
      Tab(11).Control(7)=   "lbl115"
      Tab(11).Control(8)=   "lbl118"
      Tab(11).Control(9)=   "lbl119"
      Tab(11).Control(10)=   "lbl11A"
      Tab(11).Control(11)=   "lbl11B"
      Tab(11).Control(12)=   "lbl11C"
      Tab(11).Control(13)=   "lbl11D"
      Tab(11).Control(14)=   "lbl11E"
      Tab(11).Control(15)=   "lbl11F"
      Tab(11).ControlCount=   16
      TabCaption(12)  =   "Dirs II"
      TabPicture(12)  =   "FrmMain.frx":0592
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "lbl12F"
      Tab(12).Control(1)=   "lbl12E"
      Tab(12).Control(2)=   "lbl12D"
      Tab(12).Control(3)=   "lbl12C"
      Tab(12).Control(4)=   "lbl12B"
      Tab(12).Control(5)=   "lbl12A"
      Tab(12).Control(6)=   "lbl129"
      Tab(12).Control(7)=   "lbl128"
      Tab(12).Control(8)=   "lbl127"
      Tab(12).Control(9)=   "lbl126"
      Tab(12).Control(10)=   "lbl125"
      Tab(12).Control(11)=   "lbl124"
      Tab(12).Control(12)=   "lbl123"
      Tab(12).Control(13)=   "lbl122"
      Tab(12).Control(14)=   "lbl121"
      Tab(12).Control(15)=   "lbl120"
      Tab(12).ControlCount=   16
      TabCaption(13)  =   "About"
      TabPicture(13)  =   "FrmMain.frx":05AE
      Tab(13).ControlEnabled=   0   'False
      Tab(13).Control(0)=   "Label1"
      Tab(13).Control(1)=   "cmd1"
      Tab(13).ControlCount=   2
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   3240
         Top             =   5520
      End
      Begin VB.DriveListBox Drive1 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -74880
         TabIndex        =   8
         Top             =   960
         Width           =   7215
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4650
         ItemData        =   "FrmMain.frx":05CA
         Left            =   -74880
         List            =   "FrmMain.frx":05CC
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   1920
         Width           =   7215
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Left            =   -74880
         TabIndex        =   6
         Top             =   2400
         Width           =   7215
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Left            =   -74880
         TabIndex        =   5
         Top             =   5400
         Width           =   7215
      End
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5190
         Left            =   -74880
         TabIndex        =   4
         Top             =   1320
         Width           =   7215
      End
      Begin VB.ListBox List5 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5460
         Left            =   -74880
         TabIndex        =   3
         Top             =   960
         Width           =   7215
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "Refresh Info"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74400
         TabIndex        =   2
         Top             =   6120
         Width           =   2055
      End
      Begin VB.ListBox List6 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5730
         Left            =   -74880
         TabIndex        =   1
         Top             =   840
         Width           =   7215
      End
      Begin MSComctlLib.ProgressBar PB1 
         Height          =   255
         Left            =   -74880
         TabIndex        =   9
         Top             =   1920
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar PB2 
         Height          =   255
         Left            =   -74880
         TabIndex        =   10
         Top             =   3000
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar PB3 
         Height          =   255
         Left            =   -74880
         TabIndex        =   11
         Top             =   4440
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lbl3B 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3B"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74880
         TabIndex        =   134
         Top             =   5040
         Width           =   7215
      End
      Begin VB.Label lbl3A 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3A"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   133
         Top             =   4680
         Width           =   6015
      End
      Begin VB.Label lbl39 
         BackStyle       =   0  'Transparent
         Caption         =   "Label39"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   132
         Top             =   4320
         Width           =   6015
      End
      Begin VB.Label lbl38 
         BackStyle       =   0  'Transparent
         Caption         =   "Label38"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   131
         Top             =   3960
         Width           =   6135
      End
      Begin VB.Label lbl37 
         BackStyle       =   0  'Transparent
         Caption         =   "Label37"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   130
         Top             =   3600
         Width           =   6255
      End
      Begin VB.Label lbl36 
         BackStyle       =   0  'Transparent
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   129
         Top             =   3240
         Width           =   5775
      End
      Begin VB.Label lbl35 
         BackStyle       =   0  'Transparent
         Caption         =   "Label35"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   128
         Top             =   2880
         Width           =   4575
      End
      Begin VB.Label lbl51 
         BackStyle       =   0  'Transparent
         Caption         =   "Label51"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   127
         Top             =   1200
         Width           =   4575
      End
      Begin VB.Label lbl50 
         BackStyle       =   0  'Transparent
         Caption         =   "Label50"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   126
         Top             =   840
         Width           =   7215
      End
      Begin VB.Label lbl8F 
         BackStyle       =   0  'Transparent
         Caption         =   "Label8F"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   125
         Top             =   6360
         Width           =   5415
      End
      Begin VB.Label lbl8E 
         BackStyle       =   0  'Transparent
         Caption         =   "Label8E"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   124
         Top             =   6000
         Width           =   5415
      End
      Begin VB.Label lbl8D 
         BackStyle       =   0  'Transparent
         Caption         =   "Label8D"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   123
         Top             =   5640
         Width           =   5415
      End
      Begin VB.Label lbl8C 
         BackStyle       =   0  'Transparent
         Caption         =   "Label8C"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   122
         Top             =   5280
         Width           =   5415
      End
      Begin VB.Label lbl8B 
         BackStyle       =   0  'Transparent
         Caption         =   "Label8B"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   121
         Top             =   4920
         Width           =   5415
      End
      Begin VB.Label lbl1E 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1E"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   120
         Top             =   3840
         Width           =   4575
      End
      Begin VB.Label lbl8A 
         BackStyle       =   0  'Transparent
         Caption         =   "Label8A"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   119
         Top             =   4560
         Width           =   6975
      End
      Begin VB.Label lbl1D 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1D"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   118
         Top             =   6240
         Width           =   3615
      End
      Begin VB.Label lbl1C 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1C"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   117
         Top             =   5880
         Width           =   3615
      End
      Begin VB.Label lbl60 
         BackStyle       =   0  'Transparent
         Caption         =   "Label60"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   116
         Top             =   840
         Width           =   4215
      End
      Begin VB.Label lbl12F 
         BackStyle       =   0  'Transparent
         Caption         =   "Label12F"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   115
         Top             =   6360
         Width           =   7335
      End
      Begin VB.Label lbl12E 
         BackStyle       =   0  'Transparent
         Caption         =   "Label12E"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   114
         Top             =   6000
         Width           =   7335
      End
      Begin VB.Label lbl12D 
         BackStyle       =   0  'Transparent
         Caption         =   "Label12D"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   113
         Top             =   5640
         Width           =   7335
      End
      Begin VB.Label lbl12C 
         BackStyle       =   0  'Transparent
         Caption         =   "Label12C"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   112
         Top             =   5280
         Width           =   7335
      End
      Begin VB.Label lbl12B 
         BackStyle       =   0  'Transparent
         Caption         =   "Label12B"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   111
         Top             =   4920
         Width           =   7335
      End
      Begin VB.Label lbl12A 
         BackStyle       =   0  'Transparent
         Caption         =   "Label12A"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   110
         Top             =   4560
         Width           =   7335
      End
      Begin VB.Label lbl129 
         BackStyle       =   0  'Transparent
         Caption         =   "Label129"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   109
         Top             =   4200
         Width           =   7335
      End
      Begin VB.Label lbl128 
         BackStyle       =   0  'Transparent
         Caption         =   "Label128"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   108
         Top             =   3840
         Width           =   7335
      End
      Begin VB.Label lbl127 
         BackStyle       =   0  'Transparent
         Caption         =   "Label127"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   107
         Top             =   3480
         Width           =   7335
      End
      Begin VB.Label lbl126 
         BackStyle       =   0  'Transparent
         Caption         =   "Label126"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   106
         Top             =   3120
         Width           =   7335
      End
      Begin VB.Label lbl125 
         BackStyle       =   0  'Transparent
         Caption         =   "Label125"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   105
         Top             =   2760
         Width           =   7335
      End
      Begin VB.Label lbl124 
         BackStyle       =   0  'Transparent
         Caption         =   "Label124"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   104
         Top             =   2400
         Width           =   7335
      End
      Begin VB.Label lbl123 
         BackStyle       =   0  'Transparent
         Caption         =   "Label123"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   103
         Top             =   2040
         Width           =   7335
      End
      Begin VB.Label lbl122 
         BackStyle       =   0  'Transparent
         Caption         =   "Label122"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   102
         Top             =   1680
         Width           =   7335
      End
      Begin VB.Label lbl121 
         BackStyle       =   0  'Transparent
         Caption         =   "Label121"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   101
         Top             =   1320
         Width           =   7335
      End
      Begin VB.Label lbl120 
         BackStyle       =   0  'Transparent
         Caption         =   "Label120"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   100
         Top             =   960
         Width           =   7335
      End
      Begin VB.Label lbl11F 
         BackStyle       =   0  'Transparent
         Caption         =   "Label11F"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   99
         Top             =   6360
         Width           =   7095
      End
      Begin VB.Label lbl89 
         BackStyle       =   0  'Transparent
         Caption         =   "Label89"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   98
         Top             =   4200
         Width           =   6975
      End
      Begin VB.Label lbl1B 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1B"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   97
         Top             =   5520
         Width           =   3735
      End
      Begin VB.Label lbl1A 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1A"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   96
         Top             =   5160
         Width           =   4575
      End
      Begin VB.Label lbl19 
         BackStyle       =   0  'Transparent
         Caption         =   "Label19"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   95
         Top             =   4800
         Width           =   4575
      End
      Begin VB.Label lbl18 
         BackStyle       =   0  'Transparent
         Caption         =   "Label18"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   94
         Top             =   4440
         Width           =   4575
      End
      Begin VB.Label lbl11E 
         BackStyle       =   0  'Transparent
         Caption         =   "Label11E"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   93
         Top             =   6000
         Width           =   7095
      End
      Begin VB.Label lbl11D 
         BackStyle       =   0  'Transparent
         Caption         =   "Label11D"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   92
         Top             =   5640
         Width           =   7095
      End
      Begin VB.Label lbl11C 
         BackStyle       =   0  'Transparent
         Caption         =   "Label11C"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   91
         Top             =   5280
         Width           =   7095
      End
      Begin VB.Label lbl11B 
         BackStyle       =   0  'Transparent
         Caption         =   "Label11B"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   90
         Top             =   4920
         Width           =   7095
      End
      Begin VB.Label lbl11A 
         BackStyle       =   0  'Transparent
         Caption         =   "Label11A"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   89
         Top             =   4560
         Width           =   7095
      End
      Begin VB.Label lbl119 
         BackStyle       =   0  'Transparent
         Caption         =   "Label119"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   88
         Top             =   4200
         Width           =   7095
      End
      Begin VB.Label lbl118 
         BackStyle       =   0  'Transparent
         Caption         =   "Label118"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   87
         Top             =   3840
         Width           =   7095
      End
      Begin VB.Label lbl115 
         BackStyle       =   0  'Transparent
         Caption         =   "Label115"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   86
         Top             =   2760
         Width           =   5895
      End
      Begin VB.Label lbl116 
         BackStyle       =   0  'Transparent
         Caption         =   "Label116"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   85
         Top             =   3120
         Width           =   5895
      End
      Begin VB.Label lbl117 
         BackStyle       =   0  'Transparent
         Caption         =   "Label117"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   84
         Top             =   3480
         Width           =   7215
      End
      Begin VB.Label lbl114 
         BackStyle       =   0  'Transparent
         Caption         =   "Label114"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   83
         Top             =   2400
         Width           =   5895
      End
      Begin VB.Label lbl113 
         BackStyle       =   0  'Transparent
         Caption         =   "Label113"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   82
         Top             =   2040
         Width           =   5895
      End
      Begin VB.Label lbl112 
         BackStyle       =   0  'Transparent
         Caption         =   "Label112"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   81
         Top             =   1680
         Width           =   5895
      End
      Begin VB.Label lbl111 
         BackStyle       =   0  'Transparent
         Caption         =   "Label111"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   80
         Top             =   1320
         Width           =   5895
      End
      Begin VB.Label lbl110 
         BackStyle       =   0  'Transparent
         Caption         =   "Label110"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   79
         Top             =   960
         Width           =   5895
      End
      Begin VB.Label lbl17 
         BackStyle       =   0  'Transparent
         Caption         =   "Label17"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   78
         Top             =   3480
         Width           =   4575
      End
      Begin VB.Label lbl10B 
         BackStyle       =   0  'Transparent
         Caption         =   "Label10B"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   77
         Top             =   5520
         Width           =   5655
      End
      Begin VB.Label lbl10A 
         BackStyle       =   0  'Transparent
         Caption         =   "Label10A"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   76
         Top             =   5160
         Width           =   3975
      End
      Begin VB.Label lbl109 
         BackStyle       =   0  'Transparent
         Caption         =   "Label109"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   75
         Top             =   4800
         Width           =   3975
      End
      Begin VB.Label lbl108 
         BackStyle       =   0  'Transparent
         Caption         =   "Label108"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   74
         Top             =   4440
         Width           =   3975
      End
      Begin VB.Label lbl107 
         BackStyle       =   0  'Transparent
         Caption         =   "Label107"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   73
         Top             =   4080
         Width           =   3375
      End
      Begin VB.Label lbl87 
         BackStyle       =   0  'Transparent
         Caption         =   "Label87"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   72
         Top             =   3480
         Width           =   6975
      End
      Begin VB.Label lbl86 
         BackStyle       =   0  'Transparent
         Caption         =   "Label86"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   71
         Top             =   3120
         Width           =   6975
      End
      Begin VB.Label lbl85 
         BackStyle       =   0  'Transparent
         Caption         =   "Label85"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   70
         Top             =   2760
         Width           =   6975
      End
      Begin VB.Label lbl9 
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   69
         Top             =   4320
         Width           =   6975
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl84 
         BackStyle       =   0  'Transparent
         Caption         =   "Label84"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   68
         Top             =   2400
         Width           =   6975
      End
      Begin VB.Label lbl83 
         BackStyle       =   0  'Transparent
         Caption         =   "Label83"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   67
         Top             =   2040
         Width           =   6975
      End
      Begin VB.Label lbl82 
         BackStyle       =   0  'Transparent
         Caption         =   "Label82"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   66
         Top             =   1680
         Width           =   6975
      End
      Begin VB.Label lbl81 
         BackStyle       =   0  'Transparent
         Caption         =   "Label81"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   65
         Top             =   1320
         Width           =   6975
      End
      Begin VB.Label lbl80 
         BackStyle       =   0  'Transparent
         Caption         =   "Label80"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   64
         Top             =   960
         Width           =   6975
      End
      Begin VB.Label lblD 
         BackStyle       =   0  'Transparent
         Caption         =   "LabelD"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   63
         Top             =   6240
         Width           =   7455
      End
      Begin VB.Label lblC 
         BackStyle       =   0  'Transparent
         Caption         =   "LabelC"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   62
         Top             =   5880
         Width           =   7455
      End
      Begin VB.Label lblB 
         BackStyle       =   0  'Transparent
         Caption         =   "LabelB"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   61
         Top             =   5520
         Width           =   5415
      End
      Begin VB.Label lblA 
         BackStyle       =   0  'Transparent
         Caption         =   "LabelA"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   60
         Top             =   4680
         Width           =   7095
      End
      Begin VB.Label Lbl30 
         BackStyle       =   0  'Transparent
         Caption         =   "Label30"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   59
         Top             =   1065
         Width           =   4575
      End
      Begin VB.Label Lbl31 
         BackStyle       =   0  'Transparent
         Caption         =   "Label31"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   58
         Top             =   1425
         Width           =   4575
      End
      Begin VB.Label Lbl32 
         BackStyle       =   0  'Transparent
         Caption         =   "Label32"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   57
         Top             =   1785
         Width           =   4575
      End
      Begin VB.Label Lbl33 
         BackStyle       =   0  'Transparent
         Caption         =   "Label33"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   56
         Top             =   2145
         Width           =   4575
      End
      Begin VB.Label Lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   54
         Top             =   1125
         Width           =   7095
      End
      Begin VB.Label Lbl4 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   2520
         Width           =   5055
      End
      Begin VB.Label Lbl3 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   52
         Top             =   2160
         Width           =   6735
      End
      Begin VB.Label Lbl16 
         BackStyle       =   0  'Transparent
         Caption         =   "Label16"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   51
         Top             =   960
         Width           =   4575
      End
      Begin VB.Label Lbl10 
         BackStyle       =   0  'Transparent
         Caption         =   "Label10"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   50
         Top             =   1320
         Width           =   4575
      End
      Begin VB.Label Lbl11 
         BackStyle       =   0  'Transparent
         Caption         =   "Label11"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   49
         Top             =   1680
         Width           =   7095
      End
      Begin VB.Label Lbl2 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   1800
         Width           =   6975
      End
      Begin VB.Label Lbl8 
         BackStyle       =   0  'Transparent
         Caption         =   "Label8"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   47
         Top             =   3960
         Width           =   7215
         WordWrap        =   -1  'True
      End
      Begin VB.Label Lbl7 
         BackStyle       =   0  'Transparent
         Caption         =   "Label7"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   3600
         Width           =   6975
      End
      Begin VB.Label Lbl6 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   45
         Top             =   3240
         Width           =   7095
      End
      Begin VB.Label Lbl12 
         BackStyle       =   0  'Transparent
         Caption         =   "Label12"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   44
         Top             =   2040
         Width           =   4575
      End
      Begin VB.Label Lbl15 
         BackStyle       =   0  'Transparent
         Caption         =   "Label15"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   43
         Top             =   3120
         Width           =   4575
      End
      Begin VB.Label Lbl14 
         BackStyle       =   0  'Transparent
         Caption         =   "Label14"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   42
         Top             =   2760
         Width           =   4575
      End
      Begin VB.Label Lbl13 
         BackStyle       =   0  'Transparent
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   41
         Top             =   2400
         Width           =   4575
      End
      Begin VB.Label Lbl26 
         BackStyle       =   0  'Transparent
         Caption         =   "Label26"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   40
         Top             =   4080
         Width           =   4575
      End
      Begin VB.Label Lbl25 
         BackStyle       =   0  'Transparent
         Caption         =   "Label25"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   39
         Top             =   3720
         Width           =   4575
      End
      Begin VB.Label Lbl24 
         BackStyle       =   0  'Transparent
         Caption         =   "Label24"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   38
         Top             =   3360
         Width           =   4575
      End
      Begin VB.Label Lbl23 
         BackStyle       =   0  'Transparent
         Caption         =   "Label23"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   37
         Top             =   2640
         Width           =   4575
      End
      Begin VB.Label Lbl22 
         BackStyle       =   0  'Transparent
         Caption         =   "Label22"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   36
         Top             =   2280
         Width           =   4575
      End
      Begin VB.Label Lbl21 
         BackStyle       =   0  'Transparent
         Caption         =   "Label21"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   35
         Top             =   1485
         Width           =   4575
      End
      Begin VB.Label Lbl20 
         BackStyle       =   0  'Transparent
         Caption         =   "Label20"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   34
         Top             =   1125
         Width           =   4575
      End
      Begin VB.Label Lbl28 
         BackStyle       =   0  'Transparent
         Caption         =   "Label28"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   33
         Top             =   5160
         Width           =   4095
      End
      Begin VB.Label Lbl27 
         BackStyle       =   0  'Transparent
         Caption         =   "Label27"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   32
         Top             =   4800
         Width           =   4575
      End
      Begin VB.Label Lbl34 
         BackStyle       =   0  'Transparent
         Caption         =   "Label34"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   31
         Top             =   2520
         Width           =   4575
      End
      Begin VB.Label lbl88 
         BackStyle       =   0  'Transparent
         Caption         =   "Label88"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   30
         Top             =   3840
         Width           =   6975
      End
      Begin VB.Label lbl70 
         BackStyle       =   0  'Transparent
         Caption         =   "Label70"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   29
         Top             =   960
         Width           =   4335
      End
      Begin VB.Label lbl100 
         BackStyle       =   0  'Transparent
         Caption         =   "Label100"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   28
         Top             =   1560
         Width           =   7215
      End
      Begin VB.Label lbl101 
         BackStyle       =   0  'Transparent
         Caption         =   "Label101"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   27
         Top             =   1920
         Width           =   7215
      End
      Begin VB.Label lbl102 
         BackStyle       =   0  'Transparent
         Caption         =   "Label102"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   26
         Top             =   2280
         Width           =   7215
      End
      Begin VB.Label lbl103 
         BackStyle       =   0  'Transparent
         Caption         =   "Label103"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   25
         Top             =   2640
         Width           =   2895
      End
      Begin VB.Label lbl104 
         BackStyle       =   0  'Transparent
         Caption         =   "Label104"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   24
         Top             =   3000
         Width           =   4455
      End
      Begin VB.Label lbl106 
         BackStyle       =   0  'Transparent
         Caption         =   "Label106"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   23
         Top             =   3720
         Width           =   4575
      End
      Begin VB.Label lbl105 
         BackStyle       =   0  'Transparent
         Caption         =   "Label105"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   22
         Top             =   3360
         Width           =   3135
      End
      Begin VB.Label lbl52 
         BackStyle       =   0  'Transparent
         Caption         =   "Label52"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   21
         Top             =   1560
         Width           =   4935
      End
      Begin VB.Label lbl61 
         BackStyle       =   0  'Transparent
         Caption         =   "Label61"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   20
         Top             =   1200
         Width           =   4815
      End
      Begin VB.Label lbl62 
         BackStyle       =   0  'Transparent
         Caption         =   "Label62"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   19
         Top             =   1560
         Width           =   4815
      End
      Begin VB.Label lbl53 
         BackStyle       =   0  'Transparent
         Caption         =   "Label53"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   18
         Top             =   1920
         Width           =   4935
      End
      Begin VB.Label lbl54 
         BackStyle       =   0  'Transparent
         Caption         =   "Label54"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   17
         Top             =   3960
         Width           =   4935
      End
      Begin VB.Label lbl55 
         BackStyle       =   0  'Transparent
         Caption         =   "Label55"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   16
         Top             =   4320
         Width           =   4935
      End
      Begin VB.Label lbl56 
         BackStyle       =   0  'Transparent
         Caption         =   "Label56"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   15
         Top             =   4680
         Width           =   4935
      End
      Begin VB.Label lbl57 
         BackStyle       =   0  'Transparent
         Caption         =   "Label57"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   14
         Top             =   5040
         Width           =   4935
      End
      Begin VB.Label lbl10C 
         BackStyle       =   0  'Transparent
         Caption         =   "Label10C"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   13
         Top             =   5880
         Width           =   6015
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   6135
         Left            =   -75000
         TabIndex        =   12
         Top             =   720
         Width           =   7455
      End
      Begin VB.Label Lbl5 
         BackStyle       =   0  'Transparent
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   55
         Top             =   2880
         Width           =   5415
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim prev As Boolean
Dim tret As Integer
Dim days As Integer, hours As Integer
Dim minutes As Integer, seconds As Integer, ms As Integer
Dim mLeft1 As Double, mLeft2 As Double, mLeft3 As Double

Private Sub cmd1_Click()
Form_Load
End Sub
Private Sub Drive1_Change()
DriveInfo.DiskInfo
End Sub
Private Sub Form_Load()
prev = False
If App.PrevInstance = True Then
prev = True
Unload FrmMain
Else
FrmMain.Caption = "CompInfo " + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision)
StrTmp = String(16, " ") + "CompInfo Version" + vbCrLf + String(23, " ") + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision) + vbCrLf + vbCrLf
StrTmp = StrTmp + String(10, " ") + "Programmed by Andrejevic Nemanja." + vbCrLf
StrTmp = StrTmp + String(10, " ") + "For any questions and bug reports send " + String(10, " ") + "me email on:" + vbCrLf + vbCrLf
StrTmp = StrTmp + String(13, " ") + "Hansol@Verat.net"
Label1.Caption = StrTmp
ret = GetTickCount
days = Int(ret / MPDay)
mLeft1 = ret Mod MPDay
hours = Int(mLeft1 / MPHour)
mLeft2 = mLeft1 Mod MPHour
minutes = Int(mLeft2 / MPMinute)
mLeft3 = mLeft2 Mod MPMinute
seconds = Int(mLeft3 / 1000)
ms = mLeft3 Mod 1000
FrmMain.lblD.Caption = "Time left since Windows Start-Up:  " + CStr(days) + " day(s)," + CStr(hours) + " hour(s)," + CStr(minutes) + " min(s)," + CStr(seconds) + " sec(s)"
Mod1.GetCompName
Mod1.GetUserName
BIOSInfo
CpuInfo
WinInfoDisp
WinInfoII
MouseInfo
KeyInfo
MemDisp
DriveInfo.DiskInfo
If Len(lblA.Caption) > 70 Then
    StrTmp = lblA.Caption
    lblA.Caption = Left$(StrTmp, 70) & vbCrLf
    lblA.Caption = lblA.Caption & Right$(StrTmp, Len(StrTmp) - 70)
End If
If Len(lbl3B.Caption) > 70 Then
    StrTmp = lbl3B.Caption
    lbl3B.Caption = Left$(StrTmp, 70) & vbCrLf
    lbl3B.Caption = lblA.Caption & Right$(StrTmp, Len(StrTmp) - 70)
End If
GetSysPower
SoundCard
GCard
DirsInfo.GetConfigPath
DirsInfo.GetICMPath
DirsInfo.GetMediaPath
DirsInfo.GetDevicePath
DirsInfo.GetOtherDevicePath
DirsInfo.WinDir
DirsInfo.SysDir
DirsInfo.TempDir
DirsInfo.WinBootDir
DirsInfo.GetCommonFilesPath
DirsInfo.GetProgramFilesPath
DirsInfo.GetWallPaperPath
DirsInfo.GetPersonalPath
DirsInfo.GetCommonAppDataPath
DirsInfo.GetCommonDesktopPath
DirsInfo.GetCommonStartupPath
DirsInfo.AddFolders
WinEnvironment
FontsInfo.FontsSmooth
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
If prev = False Then
    ret = MsgBox("Do you really wanna Exit?", vbQuestion + vbApplicationModal + vbOKCancel, "Exit CompInfo")
    If ret = vbOK Then
       End
    Else
        Cancel = True
    End If
End If
End Sub

Private Sub Timer1_Timer()
Mod1.MemDisp
ret = GetTickCount
days = Int(ret / MPDay)
mLeft1 = ret Mod MPDay
hours = Int(mLeft1 / MPHour)
mLeft2 = mLeft1 Mod MPHour
minutes = Int(mLeft2 / MPMinute)
mLeft3 = mLeft2 Mod MPMinute
seconds = Int(mLeft3 / 1000)
ms = mLeft3 Mod 1000
FrmMain.lblD.Caption = "Time left since Windows Start-Up:  " + CStr(days) + " day(s)," + CStr(hours) + " hour(s)," + CStr(minutes) + " min(s)," + CStr(seconds) + " sec(s)"
End Sub
